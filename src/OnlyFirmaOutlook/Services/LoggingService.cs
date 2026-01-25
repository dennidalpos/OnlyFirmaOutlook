using System.IO;
using System.Text;

namespace OnlyFirmaOutlook.Services;

/// <summary>
/// Servizio di logging centralizzato che scrive sia in memoria (per UI) che su file.
/// Thread-safe, singleton.
/// </summary>
public sealed class LoggingService : IDisposable
{
    private static readonly Lazy<LoggingService> _instance = new(() => new LoggingService());
    public static LoggingService Instance => _instance.Value;

    private readonly StringBuilder _logBuffer;
    private readonly string _logFilePath;
    private readonly object _lockObject = new();
    private StreamWriter? _fileWriter;
    private bool _disposed;

    public event EventHandler<string>? LogAdded;

    private LoggingService()
    {
        _logBuffer = new StringBuilder();

        var logFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OnlyFirmaOutlook",
            "Logs");

        Directory.CreateDirectory(logFolder);
        _logFilePath = Path.Combine(logFolder, "app.log");

        try
        {
            _fileWriter = new StreamWriter(_logFilePath, append: true, Encoding.UTF8)
            {
                AutoFlush = true
            };
        }
        catch
        {
            // Se non riusciamo a scrivere su file, continuiamo solo in memoria
            _fileWriter = null;
        }
    }

    public void Log(string message)
    {
        if (_disposed) return;

        var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
        var formattedMessage = $"[{timestamp}] {message}";

        lock (_lockObject)
        {
            _logBuffer.AppendLine(formattedMessage);

            try
            {
                _fileWriter?.WriteLine(formattedMessage);
            }
            catch
            {
                // Ignora errori di scrittura su file
            }
        }

        LogAdded?.Invoke(this, formattedMessage);
    }

    public void LogError(string message, Exception? ex = null)
    {
        var errorMessage = ex != null
            ? $"ERRORE: {message} - {ex.GetType().Name}: {ex.Message}"
            : $"ERRORE: {message}";

        Log(errorMessage);

        if (ex?.StackTrace != null)
        {
            Log($"  StackTrace: {ex.StackTrace}");
        }
    }

    public void LogWarning(string message)
    {
        Log($"AVVISO: {message}");
    }

    public string GetFullLog()
    {
        lock (_lockObject)
        {
            return _logBuffer.ToString();
        }
    }

    public void Clear()
    {
        lock (_lockObject)
        {
            _logBuffer.Clear();
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        lock (_lockObject)
        {
            _fileWriter?.Dispose();
            _fileWriter = null;
        }
    }
}
