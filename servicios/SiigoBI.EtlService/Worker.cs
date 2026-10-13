using System.Diagnostics;
using System.Text;

public sealed partial class Worker : BackgroundService
{
    private readonly ILogger<Worker> _log;
    private readonly IConfiguration _cfgRoot;
    private readonly object _stateLock = new();
    private readonly Dictionary<string, byte> _runningJobs = new(StringComparer.OrdinalIgnoreCase);

    private AppConfig _cfg = default!;
    private ServiceState _state = new();

    public Worker(ILogger<Worker> log, IConfiguration cfgRoot)
    {
        _log = log;
        _cfgRoot = cfgRoot;
    }

    public override Task StartAsync(CancellationToken cancellationToken)
    {
        _cfg = _cfgRoot.Get<AppConfig>()
            ?? throw new InvalidOperationException("No se pudo cargar appsettings.json");

        Directory.CreateDirectory(_cfg.Service.BaseDir);
        Directory.CreateDirectory(_cfg.Service.LogsDir);

        _state = ServiceState.Load(_cfg.Service.StateFile);

        ValidateConfiguration();
        foreach (var job in _cfg.Jobs.Where(j => j.Kind is EtlJobKind.Rentabilidad or EtlJobKind.Productos))
            GetJobState(job.Name).IsRunning = false;

        _log.LogInformation(
            "Servicio ETL iniciado. BaseDir={BaseDir} Jobs={JobCount}",
            _cfg.Service.BaseDir,
            _cfg.Jobs.Count);

        return base.StartAsync(cancellationToken);
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        var pollSeconds = Math.Max(5, _cfg.Service.PollSeconds);

        while (!stoppingToken.IsCancellationRequested)
        {
            try
            {
                var now = DateTimeOffset.Now;

                foreach (var job in _cfg.Jobs.Where(j => j.Enabled))
                {
                    if (!ShouldRun(job, now))
                        continue;

                    if (!TryEnterJob(job.Name))
                        continue;

                    _ = Task.Run(async () =>
                    {
                        try
                        {
                            await RunJobAsync(job, now, stoppingToken);
                        }
                        finally
                        {
                            ExitJob(job.Name);
                        }
                    }, stoppingToken);
                }
            }
            catch (OperationCanceledException) when (stoppingToken.IsCancellationRequested)
            {
                break;
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "ERROR loop principal ETL");
            }

            try
            {
                await Task.Delay(TimeSpan.FromSeconds(pollSeconds), stoppingToken);
            }
            catch (OperationCanceledException) when (stoppingToken.IsCancellationRequested)
            {
                break;
            }
        }
    }

    private void ValidateConfiguration()
    {
        if (_cfg.Jobs.Count == 0)
            throw new InvalidOperationException("No hay jobs ETL configurados.");

        foreach (var job in _cfg.Jobs)
        {
            if (job.Kind is EtlJobKind.Rentabilidad or EtlJobKind.Productos)
            {
                if (job.Enabled) ValidateRentabilidad(job);
                continue;
            }

            if (!_cfg.Sources.ContainsKey(job.Source))
                throw new InvalidOperationException(
                    $"El job '{job.Name}' referencia un Source inexistente: {job.Source}");

            if (!string.IsNullOrWhiteSpace(job.FallbackSource) &&
                !_cfg.Sources.ContainsKey(job.FallbackSource))
            {
                throw new InvalidOperationException(
                    $"El job '{job.Name}' referencia un FallbackSource inexistente: {job.FallbackSource}");
            }

            if (!_cfg.SqlTargets.ContainsKey(job.SqlTarget))
                throw new InvalidOperationException(
                    $"El job '{job.Name}' referencia un SqlTarget inexistente: {job.SqlTarget}");
        }
    }

    private bool ShouldRun(EtlJobConfig job, DateTimeOffset now)
    {
        var state = GetJobState(job.Name);

        if (job.Kind is EtlJobKind.Rentabilidad or EtlJobKind.Productos)
            return ShouldRunRentabilidad(job, state, now);

        if (state.IsRunning)
            return false;

        if (job.EveryMinutes > 0)
        {
            if (state.LastRunUtc is null)
                return true;

            return (now - state.LastRunUtc.Value) >= TimeSpan.FromMinutes(job.EveryMinutes);
        }

        if (!string.IsNullOrWhiteSpace(job.AtTime))
        {
            if (!TimeOnly.TryParse(job.AtTime, out var at))
                throw new InvalidOperationException($"Hora inválida en job {job.Name}: {job.AtTime}");

            if (job.DaysOfWeek.Count > 0)
            {
                var dow = now.DayOfWeek.ToString();
                if (!job.DaysOfWeek.Any(x =>
                        string.Equals(x, dow, StringComparison.OrdinalIgnoreCase)))
                {
                    return false;
                }
            }

            var target = new DateTimeOffset(now.Year, now.Month, now.Day, at.Hour, at.Minute, 0, now.Offset);

            if (now < target || now > target.AddMinutes(2))
                return false;

            return state.LastSuccessUtc?.Date != now.Date;
        }

        return false;
    }

    private async Task RunJobAsync(EtlJobConfig job, DateTimeOffset now, CancellationToken ct)
    {
        var state = GetJobState(job.Name);
        if (job.Kind is EtlJobKind.Rentabilidad or EtlJobKind.Productos)
        {
            var day = DateOnly.FromDateTime(RentabilidadClock.Local(now).DateTime);
            if (state.ReportAttemptDay != day) state.ReportAttempts = 0;
            state.ReportAttemptDay = day;
            state.ReportAttempts++;
        }
        state.IsRunning = true;
        state.LastRunUtc = DateTimeOffset.UtcNow;
        state.LastMessage = "En ejecución";
        SaveState();

        var logPath = MakeLogPath(job.Name, now);

        try
        {
            var target = ResolveSqlTarget(job.SqlTarget);
            if (job.Kind is EtlJobKind.Rentabilidad or EtlJobKind.Productos)
                await ExecuteRentabilidadAsync(job, target, logPath, now, ct);
            else
            {
                var source = ResolveSource(job.Source);
                await ExecuteJobWithOptionalFallbackAsync(job, source, target, logPath, ct);
            }

            state.LastSuccessUtc = DateTimeOffset.UtcNow;
            state.LastMessage = "OK";
            _log.LogInformation("ETL OK {JobName}", job.Name);
        }
        catch (Exception ex)
        {
            state.LastErrorUtc = DateTimeOffset.UtcNow;
            state.LastMessage = ex.Message;
            _log.LogError(ex, "ETL ERROR {JobName}", job.Name);
        }
        finally
        {
            state.IsRunning = false;
            SaveState();
        }
    }

    private async Task ExecuteJobWithOptionalFallbackAsync(
        EtlJobConfig job,
        SourceConfig source,
        SqlTargetConfig target,
        string logPath,
        CancellationToken ct)
    {
        try
        {
            await ExecuteJobCoreAsync(job, source, target, logPath, ct);
        }
        catch (Exception primaryEx) when (!string.IsNullOrWhiteSpace(job.FallbackSource))
        {
            var fallback = ResolveSource(job.FallbackSource!);
            var fallbackLog = Path.ChangeExtension(logPath, null) + "_fallback.log";

            _log.LogWarning(
                primaryEx,
                "Falló source primario {Source} en job {Job}. Probando fallback {Fallback}",
                job.Source,
                job.Name,
                job.FallbackSource);

            await ExecuteJobCoreAsync(job, fallback, target, fallbackLog, ct);
        }
    }

    private async Task ExecuteJobCoreAsync(
        EtlJobConfig job,
        SourceConfig source,
        SqlTargetConfig target,
        string logPath,
        CancellationToken ct)
    {
        var scriptPath = ResolveScriptPath(job);
        if (!File.Exists(scriptPath))
            throw new FileNotFoundException($"No existe el script del job '{job.Name}'", scriptPath);

        var args = BuildArgs(job, source, target, scriptPath);
        var env = BuildJobEnvironment(source, target, job);

        await RunProcessToLogAsync(_cfg.Python.Exe, args, env, job.Name, logPath, ct);
    }

    private List<string> BuildArgs(
        EtlJobConfig job,
        SourceConfig source,
        SqlTargetConfig target,
        string scriptPath)
    {
        var args = new List<string> { Quote(scriptPath) };

        switch (job.Kind)
        {
            case EtlJobKind.Movimientos:
            case EtlJobKind.MovimientosFull:
            {
                var year = job.Year ?? source.Year;

                args.Add(year.ToString());

                args.Add("--dsn");
                args.Add(source.Dsn);

                args.Add("--sql_db");
                args.Add(target.Database);

                if (job.Kind == EtlJobKind.MovimientosFull)
                {
                    args.Add("--full");
                }
                else
                {
                    args.Add("--days");
                    args.Add(job.DaysBack.ToString());
                }

                args.Add("--batch");
                args.Add(job.Batch.ToString());
                break;
            }

            case EtlJobKind.Catalogos:
            {
                var year = job.Year ?? source.Year;

                args.Add("--year");
                args.Add(year.ToString());

                args.Add("--dsn");
                args.Add(source.Dsn);

                args.Add("--sql_db");
                args.Add(target.Database);

                args.Add("--batch");
                args.Add(job.Batch.ToString());
                break;
            }

            default:
                throw new NotSupportedException($"Tipo de job no soportado: {job.Kind}");
        }

        return args;
    }

    private Dictionary<string, string> BuildJobEnvironment(
        SourceConfig source,
        SqlTargetConfig target,
        EtlJobConfig job)
    {
        var env = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var kv in _cfg.Python.Env)
            env[kv.Key] = kv.Value;

        foreach (var kv in source.Env)
            env[kv.Key] = kv.Value;

        foreach (var kv in target.Env)
            env[kv.Key] = kv.Value;

        env["SIIGOBI_SOURCE_DSN"] = source.Dsn;
        env["SIIGOBI_SOURCE_USER"] = source.User ?? "";
        env["SIIGOBI_SOURCE_PASSWORD"] = source.Password ?? "";
        env["SIIGOBI_SOURCE_ROOT"] = source.RootDir ?? "";
        env["SIIGOBI_SOURCE_YEAR"] = (job.Year ?? source.Year).ToString();
        env["SIIGOBI_JOB_NAME"] = job.Name;
        env["SIIGOBI_JOB_KIND"] = job.Kind.ToString();

        env["SIIGOBI_SQL_SERVER"] = target.Server;
        env["SIIGOBI_SQL_DATABASE"] = target.Database;
        env["SIIGOBI_SQL_TRUSTED"] = target.TrustedConnection ? "1" : "0";
        env["SIIGOBI_SQL_USER"] = target.User ?? "";
        env["SIIGOBI_SQL_PASSWORD"] = target.Password ?? "";

        return env;
    }

    private string ResolveScriptPath(EtlJobConfig job)
    {
        var relative = job.Kind == EtlJobKind.Catalogos
            ? job.CatalogosScript
            : job.MovimientosScript;

        return Path.Combine(_cfg.Service.BaseDir, relative);
    }

    private SourceConfig ResolveSource(string key) =>
        _cfg.Sources.TryGetValue(key, out var value)
            ? value
            : throw new InvalidOperationException($"Source no encontrado: {key}");

    private SqlTargetConfig ResolveSqlTarget(string key) =>
        _cfg.SqlTargets.TryGetValue(key, out var value)
            ? value
            : throw new InvalidOperationException($"SqlTarget no encontrado: {key}");

    private JobState GetJobState(string jobName)
    {
        lock (_stateLock)
        {
            return _state.GetOrCreate(jobName);
        }
    }

    private void SaveState()
    {
        lock (_stateLock)
        {
            _state.Save(_cfg.Service.StateFile);
        }
    }

    private bool TryEnterJob(string jobName)
    {
        lock (_runningJobs)
        {
            return !_runningJobs.ContainsKey(jobName) && _runningJobs.TryAdd(jobName, 0);
        }
    }

    private void ExitJob(string jobName)
    {
        lock (_runningJobs)
        {
            _runningJobs.Remove(jobName);
        }
    }

    private string MakeLogPath(string tag, DateTimeOffset now)
        => Path.Combine(_cfg.Service.LogsDir, $"{tag}_{now:yyyyMMdd_HHmmss}.log");

    private async Task RunProcessToLogAsync(
        string fileName,
        List<string> args,
        IReadOnlyDictionary<string, string> extraEnv,
        string tag,
        string logPath,
        CancellationToken ct)
    {
        var psi = new ProcessStartInfo
        {
            FileName = fileName,
            Arguments = string.Join(" ", args),
            WorkingDirectory = _cfg.Service.BaseDir,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        foreach (var kv in extraEnv)
            psi.Environment[kv.Key] = kv.Value;

        using var p = new Process { StartInfo = psi };

        Directory.CreateDirectory(Path.GetDirectoryName(logPath)!);

        await using var fs = new FileStream(logPath, FileMode.Append, FileAccess.Write, FileShare.Read);
        await using var sw = new StreamWriter(fs, new UTF8Encoding(false)) { AutoFlush = true };

        _log.LogInformation("RUN {Tag}: {File} {Args}", tag, psi.FileName, psi.Arguments);

        p.Start();

        var tOut = PumpAsync(p.StandardOutput, sw, ct);
        var tErr = PumpAsync(p.StandardError, sw, ct);

        await Task.WhenAll(tOut, tErr);
        await p.WaitForExitAsync(ct);

        if (p.ExitCode != 0)
            throw new Exception($"{fileName} falló (exit={p.ExitCode}). Revisa: {logPath}");

        _log.LogInformation("END {Tag}: exit={Exit} log={Log}", tag, p.ExitCode, logPath);
    }

    private static async Task PumpAsync(StreamReader reader, StreamWriter writer, CancellationToken ct)
    {
        while (!reader.EndOfStream && !ct.IsCancellationRequested)
        {
            var line = await reader.ReadLineAsync();
            if (line is null) break;
            await writer.WriteLineAsync(line);
        }
    }

    private static string Quote(string s)
    {
        if (string.IsNullOrWhiteSpace(s))
            return "\"\"";

        if (s.Contains(' ') || s.Contains('\t') || s.Contains('\"'))
            return "\"" + s.Replace("\"", "\\\"") + "\"";

        return s;
    }
}