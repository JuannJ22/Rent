using System.Text.Json;

public sealed class ServiceState
{
    public Dictionary<string, JobState> Jobs { get; set; } = new(StringComparer.OrdinalIgnoreCase);

    public JobState GetOrCreate(string jobName)
    {
        if (!Jobs.TryGetValue(jobName, out var state))
        {
            state = new JobState();
            Jobs[jobName] = state;
        }

        return state;
    }

    public static ServiceState Load(string path)
    {
        try
        {
            if (!File.Exists(path))
                return new ServiceState();

            var json = File.ReadAllText(path);
            return JsonSerializer.Deserialize<ServiceState>(json) ?? new ServiceState();
        }
        catch
        {
            return new ServiceState();
        }
    }

    public void Save(string path)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);

        var json = JsonSerializer.Serialize(this, new JsonSerializerOptions
        {
            WriteIndented = true
        });

        File.WriteAllText(path, json);
    }
}

public sealed class JobState
{
    public DateTimeOffset? LastRunUtc { get; set; }
    public DateTimeOffset? LastSuccessUtc { get; set; }
    public DateTimeOffset? LastErrorUtc { get; set; }
    public bool IsRunning { get; set; }
    public DateOnly? ReportAttemptDay { get; set; }
    public int ReportAttempts { get; set; }
    public string? LastMessage { get; set; }
}