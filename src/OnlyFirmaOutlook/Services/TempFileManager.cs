using System.IO;

namespace OnlyFirmaOutlook.Services;

/// <summary>
/// Gestisce i file temporanei per la sessione corrente.
/// Copia i file dalla share di rete a una cartella locale temporanea.
/// Esegue cleanup best-effort alla chiusura.
/// </summary>
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

    /// <summary>
    /// Copia un file dalla sorgente (possibilmente share di rete) alla cartella temporanea locale.
    /// Restituisce il percorso del file copiato.
    /// </summary>
    public string CopyToLocalTemp(string sourceFilePath)
    {
        if (!File.Exists(sourceFilePath))
        {
            throw new FileNotFoundException($"File sorgente non trovato: {sourceFilePath}");
        }

        var fileName = Path.GetFileName(sourceFilePath);
        var destPath = Path.Combine(SessionTempFolder, fileName);

        // Se esiste già un file con lo stesso nome, aggiungi un GUID
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

    /// <summary>
    /// Verifica se un percorso è su una share di rete (UNC path).
    /// </summary>
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
            // Fallback per path che iniziano con \\
            return path.StartsWith(@"\\", StringComparison.Ordinal);
        }
    }

    /// <summary>
    /// Cleanup best-effort della cartella temporanea della sessione.
    /// Non blocca l'uscita dell'app, logga eventuali errori.
    /// </summary>
    public void CleanupSessionFolder()
    {
        if (!Directory.Exists(SessionTempFolder))
        {
            _logger.Log("Cartella temporanea sessione già eliminata o non esistente");
            return;
        }

        _logger.Log($"Avvio cleanup cartella temporanea: {SessionTempFolder}");

        var maxRetries = 3;
        var retryDelayMs = 100;

        for (var attempt = 1; attempt <= maxRetries; attempt++)
        {
            try
            {
                // Prima prova a eliminare tutti i file
                foreach (var file in Directory.GetFiles(SessionTempFolder, "*", SearchOption.AllDirectories))
                {
                    try
                    {
                        File.SetAttributes(file, FileAttributes.Normal);
                        File.Delete(file);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning($"Impossibile eliminare file '{file}': {ex.Message}");
                    }
                }

                // Poi elimina le sottocartelle
                foreach (var dir in Directory.GetDirectories(SessionTempFolder, "*", SearchOption.AllDirectories)
                                             .OrderByDescending(d => d.Length))
                {
                    try
                    {
                        Directory.Delete(dir, recursive: false);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning($"Impossibile eliminare cartella '{dir}': {ex.Message}");
                    }
                }

                // Infine elimina la cartella principale
                Directory.Delete(SessionTempFolder, recursive: true);
                _logger.Log("Cleanup cartella temporanea completato con successo");
                return;
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Tentativo {attempt}/{maxRetries} cleanup fallito: {ex.Message}");

                if (attempt < maxRetries)
                {
                    Thread.Sleep(retryDelayMs * attempt);
                }
            }
        }

        _logger.LogWarning($"Cleanup cartella temporanea non completato dopo {maxRetries} tentativi. " +
                          $"La cartella '{SessionTempFolder}' potrebbe richiedere pulizia manuale.");
    }

    /// <summary>
    /// Pulisce le cartelle temporanee orfane di sessioni precedenti.
    /// Chiamato all'avvio per manutenzione.
    /// </summary>
    public void CleanupOrphanedFolders()
    {
        if (!Directory.Exists(_baseTempFolder)) return;

        _logger.Log("Verifica cartelle temporanee orfane...");

        try
        {
            foreach (var dir in Directory.GetDirectories(_baseTempFolder))
            {
                var dirName = Path.GetFileName(dir);

                // Salta la cartella della sessione corrente
                if (dirName == _sessionId.ToString()) continue;

                // Prova a eliminare solo se la cartella è vecchia (più di 1 giorno)
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
