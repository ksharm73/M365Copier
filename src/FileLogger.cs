namespace OneDriveCopier;

/// <summary>
/// Thread-safe logger that writes to the console AND a date-stamped log file.
/// Log file name format:  OneDriveCopier_yyyyMMdd_HHmmss.log
/// Each run gets its own log file — safe for Autosys parallel scheduling.
/// </summary>
internal sealed class FileLogger : IDisposable
{
    private readonly StreamWriter _writer;
    private readonly object _lock = new();
    private bool _disposed;

    public string LogFilePath { get; }

    public FileLogger(string logDirectory)
    {
        Directory.CreateDirectory(logDirectory);

        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        LogFilePath = Path.Combine(logDirectory, $"OneDriveCopier_{timestamp}.log");

        _writer = new StreamWriter(LogFilePath, append: false, System.Text.Encoding.UTF8)
        {
            AutoFlush = true
        };

        WriteHeader();
    }

    // ── Public logging methods ──────────────────────────────────────────────

    public void Info(string message)    => Write("INFO ", message, ConsoleColor.Cyan);
    public void Success(string message) => Write("OK   ", message, ConsoleColor.Green);
    public void Warn(string message)    => Write("WARN ", message, ConsoleColor.Yellow);
    public void Error(string message)   => Write("ERROR", message, ConsoleColor.Red);

    public void Separator() => Write(null, new string('─', 72), ConsoleColor.DarkGray);

    public void Summary(int total, int copied, int skipped, int failed)
    {
        Separator();
        Info($"Total files found : {total}");
        Success($"Copied            : {copied}");
        if (skipped > 0) Warn($"Skipped (exist)   : {skipped}");
        if (failed  > 0) Error($"Failed            : {failed}");
        Separator();
    }

    // ── Private helpers ─────────────────────────────────────────────────────

    private void WriteHeader()
    {
        string header = $"""
            ══════════════════════════════════════════════════════════════════════
              OneDrive → Network Drive File Copier
              Run started : {DateTime.Now:yyyy-MM-dd HH:mm:ss}
              Machine     : {Environment.MachineName}
              User        : {Environment.UserName}
              Log file    : {LogFilePath}
            ══════════════════════════════════════════════════════════════════════
            """;

        lock (_lock)
        {
            Console.WriteLine(header);
            _writer.WriteLine(header);
        }
    }

    private void Write(string? level, string message, ConsoleColor color)
    {
        string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
        string prefix    = level is null ? "         " : $"[{level}]";
        string line      = $"{timestamp} {prefix} {message}";

        lock (_lock)
        {
            // Console output (coloured)
            Console.ForegroundColor = color;
            Console.WriteLine(line);
            Console.ResetColor();

            // File output (plain text)
            if (!_disposed)
                _writer.WriteLine(line);
        }
    }

    public void Dispose()
    {
        lock (_lock)
        {
            if (_disposed) return;
            _disposed = true;
            _writer.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} [INFO ] Run ended.");
            _writer.Dispose();
        }
    }
}
