using System.Text;

namespace TeamsMessageFetcher;

/// <summary>
/// Minimal file logger: one file per day (dailyoverview-yyyyMMdd.log).
/// Level filtering is done by the logging pipeline.
/// </summary>
public sealed class FileLoggerProvider : ILoggerProvider
{
    private readonly string _directory;
    private readonly object _lock = new();

    public FileLoggerProvider(string directory)
    {
        _directory = directory;
        Directory.CreateDirectory(_directory);
    }

    public ILogger CreateLogger(string categoryName) => new FileLogger(categoryName, this);

    public void Dispose() { }

    private void Write(string line)
    {
        var path = Path.Combine(_directory, $"dailyoverview-{DateTime.Now:yyyyMMdd}.log");
        lock (_lock)
        {
            try { File.AppendAllText(path, line + Environment.NewLine, Encoding.UTF8); }
            catch (IOException) { /* logging must never break the app */ }
            catch (UnauthorizedAccessException) { }
        }
    }

    private sealed class FileLogger(string category, FileLoggerProvider provider) : ILogger
    {
        public IDisposable? BeginScope<TState>(TState state) where TState : notnull => null;

        public bool IsEnabled(LogLevel logLevel) => logLevel != LogLevel.None;

        public void Log<TState>(LogLevel logLevel, EventId eventId, TState state, Exception? exception,
                                Func<TState, Exception?, string> formatter)
        {
            if (!IsEnabled(logLevel)) return;

            var text = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} [{ShortLevel(logLevel)}] {category}: {formatter(state, exception)}";
            if (exception != null) text += Environment.NewLine + exception;
            provider.Write(text);
        }

        private static string ShortLevel(LogLevel level) => level switch
        {
            LogLevel.Trace       => "TRC",
            LogLevel.Debug       => "DBG",
            LogLevel.Information => "INF",
            LogLevel.Warning     => "WRN",
            LogLevel.Error       => "ERR",
            LogLevel.Critical    => "CRT",
            _                    => "???",
        };
    }
}
