public sealed class AppConfig
{
    public ServiceConfig Service { get; set; } = new();
    public PythonConfig Python { get; set; } = new();
    public Dictionary<string, SourceConfig> Sources { get; set; } = new();
    public Dictionary<string, SqlTargetConfig> SqlTargets { get; set; } = new();
    public List<EtlJobConfig> Jobs { get; set; } = new();
}

public sealed class ServiceConfig
{
    public string BaseDir { get; set; } = "";
    public string LogsDir { get; set; } = "";
    public string StateFile { get; set; } = "";
    public int PollSeconds { get; set; } = 15;
}

public sealed class PythonConfig
{
    public string Exe { get; set; } = "";
    public Dictionary<string, string> Env { get; set; } = new();
}

public sealed class SourceConfig
{
    public string Type { get; set; } = "Dsn";
    public string Dsn { get; set; } = "";
    public string User { get; set; } = "";
    public string Password { get; set; } = "";
    public string RootDir { get; set; } = "";
    public int Year { get; set; }
    public Dictionary<string, string> Env { get; set; } = new();
}

public sealed class SqlTargetConfig
{
    public string Server { get; set; } = "";
    public string Database { get; set; } = "";
    public bool TrustedConnection { get; set; } = false;
    public string User { get; set; } = "";
    public string Password { get; set; } = "";
    public Dictionary<string, string> Env { get; set; } = new();
}

public enum EtlJobKind
{
    Movimientos,
    MovimientosFull,
    Catalogos,
    Rentabilidad,
    Productos
}

public sealed class EtlJobConfig
{
    public string Name { get; set; } = "";
    public bool Enabled { get; set; } = true;
    public EtlJobKind Kind { get; set; } = EtlJobKind.Movimientos;

    public RentabilidadConfig Rentabilidad { get; set; } = new();

    public string Source { get; set; } = "";
    public string? FallbackSource { get; set; }

    public string SqlTarget { get; set; } = "";

    public int? Year { get; set; }

    public int EveryMinutes { get; set; } = 0;
    public string AtTime { get; set; } = "";
    public List<string> DaysOfWeek { get; set; } = new();

    public int DaysBack { get; set; } = 3;
    public int Batch { get; set; } = 20000;

    public string MovimientosScript { get; set; } = @"scripts\etl_movimientos.py";
    public string CatalogosScript { get; set; } = @"scripts\run_all_catalogs.py";
}
public sealed class RentabilidadConfig
{
    public string PythonExe { get; set; } = @"C:\Rentabilidad\Rent\.venv-sql\Scripts\python.exe";
    public string Script { get; set; } = @"C:\Rentabilidad\Rent\servicios\etl_rentabilidad.py";
    public string ProjectDir { get; set; } = @"C:\Rentabilidad\Rent";
    public string BaseDir { get; set; } = @"C:\Rentabilidad";
    public string Template { get; set; } = @"C:\Rentabilidad\PLANTILLA.xlsx";
    public string FirstReportDate { get; set; } = "";
    public int TimeoutSeconds { get; set; } = 1800;
    public int RetryMinutes { get; set; } = 15;
    public int MaxAttempts { get; set; } = 3;
    public List<string> DependsOn { get; set; } = new();
}
