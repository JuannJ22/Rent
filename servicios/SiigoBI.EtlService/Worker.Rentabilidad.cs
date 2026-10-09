using System.Diagnostics;
using System.Globalization;
using System.Text;

public static class RentabilidadClock
{
    // Colombia usa UTC-5 todo el año; independiente de la zona del servidor.
    public static DateTimeOffset Local(DateTimeOffset value) => value.ToOffset(TimeSpan.FromHours(-5));

    public static bool IsDue(EtlJobConfig job, JobState state, DateTimeOffset now)
    {
        var local = Local(now);
        var day = DateOnly.FromDateTime(local.DateTime);
        if (job.Kind == EtlJobKind.Rentabilidad && DateOnly.TryParseExact(job.Rentabilidad.FirstReportDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var first) && day.AddDays(-1) < first)
            return false;
        if (state.IsRunning || !TimeOnly.TryParse(job.AtTime, CultureInfo.InvariantCulture, out var at)) return false;
        if (TimeOnly.FromDateTime(local.DateTime) < at) return false;
        if (state.LastSuccessUtc is { } success && DateOnly.FromDateTime(Local(success).DateTime) == day) return false;
        if (state.ReportAttemptDay == day && state.ReportAttempts >= job.Rentabilidad.MaxAttempts) return false;
        if (state.LastRunUtc is { } run && now - run < TimeSpan.FromMinutes(job.Rentabilidad.RetryMinutes)) return false;
        return true;
    }
}

public sealed partial class Worker
{
    private void ValidateRentabilidad(EtlJobConfig job)
    {
        var options = job.Rentabilidad;
        if (!string.IsNullOrWhiteSpace(options.FirstReportDate) && !DateOnly.TryParseExact(options.FirstReportDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out _))
            throw new InvalidOperationException("FirstReportDate requiere YYYY-MM-DD.");
        if (!_cfg.SqlTargets.ContainsKey(job.SqlTarget)) throw new InvalidOperationException($"SqlTarget inexistente: {job.SqlTarget}");
        if (!TimeOnly.TryParse(job.AtTime, CultureInfo.InvariantCulture, out _) || job.EveryMinutes != 0)
            throw new InvalidOperationException("Rentabilidad requiere AtTime y EveryMinutes=0.");
        if (options.TimeoutSeconds <= 0 || options.RetryMinutes <= 0 || options.MaxAttempts <= 0)
            throw new InvalidOperationException("Los limites de rentabilidad deben ser positivos.");
        foreach (var path in new[] { options.PythonExe, options.Script, options.Template })
            if (!File.Exists(path)) throw new FileNotFoundException("Archivo requerido para rentabilidad", path);
        if (!Directory.Exists(options.ProjectDir)) throw new DirectoryNotFoundException(options.ProjectDir);
        foreach (var name in options.DependsOn)
            if (!_cfg.Jobs.Any(j => (!job.Enabled || j.Enabled) && j.Name.Equals(name, StringComparison.OrdinalIgnoreCase) && j.Kind != EtlJobKind.Rentabilidad && !j.Name.Equals(job.Name, StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException($"Dependencia ETL inexistente o deshabilitada: {name}");
    }

    private bool ShouldRunRentabilidad(EtlJobConfig job, JobState state, DateTimeOffset now)
    {
        if (!RentabilidadClock.IsDue(job, state, now)) return false;
        var requiredDay = RentabilidadClock.Local(now).Date.AddDays(job.Kind == EtlJobKind.Productos ? 0 : -1);
        lock (_stateLock)
        {
            foreach (var name in job.Rentabilidad.DependsOn)
            {
                if (!_state.Jobs.TryGetValue(name, out var dependency) || dependency.IsRunning ||
                    dependency.LastSuccessUtc is not { } success || RentabilidadClock.Local(success).Date < requiredDay ||
                    (dependency.LastErrorUtc is { } error && error > success))
                    return false;
                var prerequisite = _cfg.Jobs.First(j => j.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
                if (TimeOnly.TryParse(prerequisite.AtTime, CultureInfo.InvariantCulture, out var scheduled) &&
                    RentabilidadClock.Local(success).DateTime < requiredDay.Add(scheduled.ToTimeSpan()))
                    return false;
            }
        }
        return true;
    }

    private async Task ExecuteRentabilidadAsync(EtlJobConfig job, SqlTargetConfig target, string logPath, DateTimeOffset now, CancellationToken ct)
    {
        var options = job.Rentabilidad;
        var date = RentabilidadClock.Local(now).Date.AddDays(job.Kind == EtlJobKind.Productos ? 0 : -1).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        var psi = new ProcessStartInfo(options.PythonExe)
        {
            WorkingDirectory = options.ProjectDir, UseShellExecute = false, CreateNoWindow = true,
            RedirectStandardOutput = true, RedirectStandardError = true,
            StandardOutputEncoding = Encoding.UTF8, StandardErrorEncoding = Encoding.UTF8
        };
        foreach (var value in new[] { options.Script, "--fecha", date, "--base-dir", options.BaseDir,
            "--template", options.Template, "--timeout", options.TimeoutSeconds.ToString(CultureInfo.InvariantCulture) })
            psi.ArgumentList.Add(value);
        if (job.Kind == EtlJobKind.Productos) psi.ArgumentList.Add("--capture-products");
        foreach (var item in target.Env) psi.Environment[item.Key] = item.Value;
        // Credenciales del servicio: no depende del XML DPAPI del administrador.
        psi.Environment["SQL_SERVER"] = target.Server;
        psi.Environment["SQL_DATABASE"] = "SiigoRent";
        psi.Environment["SQL_USER"] = target.User;
        psi.Environment["SQL_PASSWORD"] = target.Password;
        psi.Environment["SQL_TRUSTED"] = target.TrustedConnection ? "1" : "0";
        psi.Environment["SQL_ENCRYPT"] = "1";
        psi.Environment["SQL_TRUST_CERT"] = "0";
        psi.Environment["SQL_DRIVER"] = "ODBC Driver 18 for SQL Server";
        psi.Environment["PYTHONUTF8"] = "1";
        psi.Environment["PYTHONIOENCODING"] = "utf-8";
        psi.Environment["RENT_DIR"] = options.BaseDir;
        Directory.CreateDirectory(Path.GetDirectoryName(logPath)!);
        await using var writer = new StreamWriter(logPath, append: true, new UTF8Encoding(false)) { AutoFlush = true };
        using var gate = new SemaphoreSlim(1);
        async Task Pump(StreamReader reader)
        {
            while (await reader.ReadLineAsync() is { } line)
            {
                await gate.WaitAsync();
                try { await writer.WriteLineAsync(line); }
                finally { gate.Release(); }
            }
        }
        using var process = new Process { StartInfo = psi };
        process.Start();
        using var stopRegistration = ct.Register(() =>
        {
            try { if (!process.HasExited) process.Kill(entireProcessTree: true); }
            catch (Exception ex) { _log.LogWarning(ex, "No se pudo terminar el proceso de rentabilidad al detener el servicio."); }
        });
        var stdout = Pump(process.StandardOutput);
        var stderr = Pump(process.StandardError);
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct);
        timeout.CancelAfter(TimeSpan.FromSeconds(options.TimeoutSeconds + 30));
        try { await process.WaitForExitAsync(timeout.Token); }
        catch (OperationCanceledException)
        {
            if (!process.HasExited) process.Kill(entireProcessTree: true);
            await process.WaitForExitAsync();
            await Task.WhenAll(stdout, stderr);
            throw;
        }
        await Task.WhenAll(stdout, stderr);
        if (process.ExitCode != 0) throw new InvalidOperationException($"Rentabilidad fallo (exit={process.ExitCode}). Revisa {logPath}");
    }
}
