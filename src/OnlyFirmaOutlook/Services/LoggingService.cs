using System.IO;
using System.Text;

namespace OnlyFirmaOutlook.Services;





public sealed class LoggingService : IDisposable
{
    private static readonly Lazy<LoggingService> _instance = new(() => new LoggingService());
    public static LoggingService Instance => _instance.Value;

    private readonly string _logFilePath;
    private readonly object _lockObject = new();
    private readonly Queue<string> _logLines;
    private StreamWriter? _fileWriter;
    private bool _disposed;
    private bool _fileWriterFailureLogged;

    private const int MaxBufferedLines = 2000;
    private const long MaxLogFileSizeBytes = 5 * 1024 * 1024;

    public event EventHandler<string>? LogAdded;

    private LoggingService()
    {
        _logLines = new Queue<string>();

        var logFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OnlyFirmaOutlook",
            "Logs");

        Directory.CreateDirectory(logFolder);
        _logFilePath = Path.Combine(logFolder, "app.log");

        try
        {
            _fileWriter = CreateWriter();
        }
        catch (Exception ex)
        {
            _fileWriter = null;
            AddInternalWarning($"Logging su file non disponibile: {ex.Message}");
        }
    }

    public void Log(string message)
    {
        if (_disposed) return;

        var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
        var formattedMessage = $"[{timestamp}] INFO: {message}";

        lock (_lockObject)
        {
            EnqueueLine(formattedMessage);
            WriteToFile(formattedMessage);
        }

        LogAdded?.Invoke(this, formattedMessage);
    }

    public void LogError(string message, Exception? ex = null)
    {
        var errorMessage = ex != null
            ? $"{message} - {ex.GetType().Name}: {ex.Message}"
            : message;

        LogWithLevel("ERROR", errorMessage);

        if (ex?.StackTrace != null)
        {
            Log($"  StackTrace: {ex.StackTrace}");
        }
    }

    public void LogWarning(string message)
    {
        LogWithLevel("WARN", message);
    }

    private void LogWithLevel(string level, string message)
    {
        if (_disposed) return;

        var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
        var formattedMessage = $"[{timestamp}] {level}: {message}";

        lock (_lockObject)
        {
            EnqueueLine(formattedMessage);
            WriteToFile(formattedMessage);
        }

        LogAdded?.Invoke(this, formattedMessage);
    }

    public string GetFullLog()
    {
        lock (_lockObject)
        {
            return string.Join(Environment.NewLine, _logLines);
        }
    }

    public void Clear()
    {
        lock (_lockObject)
        {
            _logLines.Clear();
            _fileWriter?.Dispose();
            _fileWriter = null;

            try
            {
                if (File.Exists(_logFilePath))
                {
                    File.Delete(_logFilePath);
                }

                _fileWriterFailureLogged = false;
                _fileWriter = CreateWriter();
            }
            catch (Exception ex)
            {
                _fileWriter = null;
                AddInternalWarning($"Pulizia log fallita: {ex.Message}");
            }
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

    private void EnqueueLine(string line)
    {
        _logLines.Enqueue(line);
        while (_logLines.Count > MaxBufferedLines)
        {
            _logLines.Dequeue();
        }
    }

    private StreamWriter CreateWriter()
    {
        return new StreamWriter(_logFilePath, append: true, Encoding.UTF8)
        {
            AutoFlush = true
        };
    }

    private void WriteToFile(string message)
    {
        if (_fileWriter == null)
        {
            return;
        }

        try
        {
            _fileWriter.WriteLine(message);
            RotateIfNeeded();
        }
        catch (Exception ex)
        {
            _fileWriter?.Dispose();
            _fileWriter = null;
            AddInternalWarning($"Errore scrittura log su file: {ex.Message}");
        }
    }

    private void RotateIfNeeded()
    {
        try
        {
            var fileInfo = new FileInfo(_logFilePath);
            if (!fileInfo.Exists || fileInfo.Length < MaxLogFileSizeBytes)
            {
                return;
            }

            _fileWriter?.Dispose();
            _fileWriter = null;

            var rotatedPath = _logFilePath + ".1";
            if (File.Exists(rotatedPath))
            {
                File.Delete(rotatedPath);
            }

            File.Move(_logFilePath, rotatedPath);
            _fileWriter = CreateWriter();
        }
        catch (Exception ex)
        {
            AddInternalWarning($"Rotazione log fallita: {ex.Message}");
        }
    }

    private void AddInternalWarning(string message)
    {
        if (_fileWriterFailureLogged)
        {
            return;
        }

        _fileWriterFailureLogged = true;
        var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
        var formattedMessage = $"[{timestamp}] WARN: {message}";

        lock (_lockObject)
        {
            EnqueueLine(formattedMessage);
        }

        LogAdded?.Invoke(this, formattedMessage);
    }
}
