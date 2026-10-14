using System.Runtime.CompilerServices;
using Microsoft.Extensions.Logging.Abstractions;

namespace TeamsMessageFetcher;

/// <summary>
/// Central entry point for logging from fetchers and pages.
/// The category is derived from the calling file name.
/// </summary>
public static class FetchLog
{
    private static ILoggerFactory _factory = NullLoggerFactory.Instance;

    public static void Initialize(ILoggerFactory factory) => _factory = factory;

    public static void Debug(string message, [CallerFilePath] string file = "")
        => Write(LogLevel.Debug, message, null, file);

    public static void Info(string message, [CallerFilePath] string file = "")
        => Write(LogLevel.Information, message, null, file);

    public static void Warning(string message, [CallerFilePath] string file = "")
        => Write(LogLevel.Warning, message, null, file);

    public static void Error(string message, Exception? exception = null, [CallerFilePath] string file = "")
        => Write(LogLevel.Error, message, exception, file);

    private static void Write(LogLevel level, string message, Exception? exception, string file)
        => _factory.CreateLogger(Path.GetFileNameWithoutExtension(file)).Log(level, exception, "{Message}", message);
}
