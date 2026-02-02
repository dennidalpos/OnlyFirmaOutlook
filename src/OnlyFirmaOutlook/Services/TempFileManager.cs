using System.IO;

namespace OnlyFirmaOutlook.Services;






public sealed class TempFileManager
{
    private static readonly Lazy<TempFileManager> _instance = new(() => new TempFileManager());
    public static TempFileManager Instance => _instance.Value;

    private readonly LoggingService _logger;
    private readonly string _baseTempFolder;
    private readonly Guid _sessionId;

    public string SessionTempFolder { get; }

    private TempFileManager()
    {
        _logger = LoggingService.Instance;
        _sessionId = Guid.NewGuid();

        _baseTempFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OnlyFirmaOutlook",
            "Temp");

        SessionTempFolder = Path.Combine(_baseTempFolder, _sessionId.ToString());

        try
        {
            Directory.CreateDirectory(SessionTempFolder);
            _logger.Log($"Cartella temporanea sessione creata: {SessionTempFolder}");
        }
        catch (Exception ex)
        {
            _logger.LogError("Impossibile creare cartella temporanea sessione", ex);
            throw;
        }
    }

    
    
    
    
    public string CopyToLocalTemp(string sourceFilePath)
    {
        if (!File.Exists(sourceFilePath))
        {
            throw new FileNotFoundException($"File sorgente non trovato: {sourceFilePath}");
        }

        var fileName = Path.GetFileName(sourceFilePath);
        var destPath = Path.Combine(SessionTempFolder, fileName);

        
        if (File.Exists(destPath))
        {
            var nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
            var ext = Path.GetExtension(fileName);
            destPath = Path.Combine(SessionTempFolder, $"{nameWithoutExt}_{Guid.NewGuid():N}{ext}");
        }

        _logger.Log($"Copia file da '{sourceFilePath}' a '{destPath}'");

        try
        {
            File.Copy(sourceFilePath, destPath, overwrite: true);
            _logger.Log("File copiato con successo");
            return destPath;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Errore durante la copia del file", ex);
            throw;
        }
    }

    
    
    
    public static bool IsUncPath(string path)
    {
        if (string.IsNullOrEmpty(path)) return false;

        try
        {
            var uri = new Uri(path);
            return uri.IsUnc;
        }
        catch
        {
            
            return path.StartsWith(@"\\", StringComparison.Ordinal);
        }
    }

    
    
    
    
    public void CleanupSessionFolder()
    {
        if (!Directory.Exists(SessionTempFolder))
        {
            _logger.Log("Cartella temporanea sessione già eliminata o non esistente");
            return;
        }

        TempCleanupHelper.CleanupDirectoryWithRetries(
            SessionTempFolder,
            _logger,
            "cartella temporanea sessione");
    }

    
    
    
    
    public void CleanupOrphanedFolders()
    {
        if (!Directory.Exists(_baseTempFolder)) return;

        _logger.Log("Verifica cartelle temporanee orfane...");

        try
        {
            foreach (var dir in Directory.GetDirectories(_baseTempFolder))
            {
                var dirName = Path.GetFileName(dir);

                
                if (dirName == _sessionId.ToString()) continue;

                
                var dirInfo = new DirectoryInfo(dir);
                if (dirInfo.LastWriteTime < DateTime.Now.AddDays(-1))
                {
                    try
                    {
                        Directory.Delete(dir, recursive: true);
                        _logger.Log($"Eliminata cartella temporanea orfana: {dir}");
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning($"Impossibile eliminare cartella orfana '{dir}': {ex.Message}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Errore durante pulizia cartelle orfane: {ex.Message}");
        }
    }
}
