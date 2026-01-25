using System.IO;
using OnlyFirmaOutlook.Models;

namespace OnlyFirmaOutlook.Services;

/// <summary>
/// Gestisce le firme esistenti nella cartella di destinazione.
/// Permette di elencare ed eliminare le firme.
/// </summary>
public class SignatureRepository
{
    private readonly LoggingService _logger;

    public SignatureRepository()
    {
        _logger = LoggingService.Instance;
    }

    /// <summary>
    /// Ottiene la cartella firme predefinita di Outlook.
    /// </summary>
    public static string GetDefaultOutlookSignaturesFolder()
    {
        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Microsoft",
            "Signatures");
    }

    /// <summary>
    /// Ottiene la cartella di output alternativa quando Outlook non è disponibile.
    /// </summary>
    public static string GetAlternativeOutputFolder()
    {
        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "OnlyFirmaOutlook",
            "Output");
    }

    /// <summary>
    /// Elenca tutte le firme presenti nella cartella specificata.
    /// Una firma è identificata dalla presenza di un file .htm.
    /// </summary>
    public List<SignatureInfo> GetSignatures(string folderPath)
    {
        var signatures = new List<SignatureInfo>();

        if (!Directory.Exists(folderPath))
        {
            _logger.Log($"Cartella firme non esistente: {folderPath}");
            return signatures;
        }

        _logger.Log($"Ricerca firme in: {folderPath}");

        try
        {
            var htmFiles = Directory.GetFiles(folderPath, "*.htm", SearchOption.TopDirectoryOnly);

            foreach (var htmFile in htmFiles)
            {
                var baseName = Path.GetFileNameWithoutExtension(htmFile);
                var signature = new SignatureInfo
                {
                    Name = baseName,
                    FolderPath = folderPath,
                    HasHtm = true,
                    HasRtf = File.Exists(Path.Combine(folderPath, baseName + ".rtf")),
                    HasTxt = File.Exists(Path.Combine(folderPath, baseName + ".txt")),
                    HasFilesFolder = Directory.Exists(Path.Combine(folderPath, baseName + "_files")),
                    HasFileFolder = Directory.Exists(Path.Combine(folderPath, baseName + "_file"))
                };

                signatures.Add(signature);
                _logger.Log($"Trovata firma: {baseName}");
            }

            _logger.Log($"Totale firme trovate: {signatures.Count}");
        }
        catch (Exception ex)
        {
            _logger.LogError($"Errore durante lettura firme", ex);
        }

        return signatures.OrderBy(s => s.Name).ToList();
    }

    /// <summary>
    /// Elimina una firma e tutti i suoi file associati.
    /// </summary>
    public bool DeleteSignature(SignatureInfo signature)
    {
        if (signature == null)
        {
            _logger.LogError("Tentativo di eliminare firma null");
            return false;
        }

        _logger.Log($"Eliminazione firma: {signature.Name}");

        var success = true;
        var basePath = Path.Combine(signature.FolderPath, signature.Name);

        // Elimina file .htm
        success &= TryDeleteFile(basePath + ".htm");

        // Elimina file .rtf
        success &= TryDeleteFile(basePath + ".rtf");

        // Elimina file .txt
        success &= TryDeleteFile(basePath + ".txt");

        // Elimina cartella _files
        success &= TryDeleteDirectory(basePath + "_files");

        // Elimina cartella _file
        success &= TryDeleteDirectory(basePath + "_file");

        if (success)
        {
            _logger.Log($"Firma '{signature.Name}' eliminata con successo");
        }
        else
        {
            _logger.LogWarning($"Eliminazione firma '{signature.Name}' completata con alcuni errori");
        }

        return success;
    }

    /// <summary>
    /// Elimina i file esistenti di una firma prima di sovrascriverla.
    /// </summary>
    public void DeleteExistingSignatureFiles(string folderPath, string signatureName)
    {
        _logger.Log($"Eliminazione file firma esistente: {signatureName}");

        var basePath = Path.Combine(folderPath, signatureName);

        TryDeleteFile(basePath + ".htm");
        TryDeleteFile(basePath + ".rtf");
        TryDeleteFile(basePath + ".txt");
        TryDeleteDirectory(basePath + "_files");
        TryDeleteDirectory(basePath + "_file");
    }

    /// <summary>
    /// Verifica se una firma con il nome specificato esiste già.
    /// </summary>
    public bool SignatureExists(string folderPath, string signatureName)
    {
        var htmPath = Path.Combine(folderPath, signatureName + ".htm");
        return File.Exists(htmPath);
    }

    /// <summary>
    /// Verifica se è possibile scrivere nella cartella specificata.
    /// Crea e elimina un file temporaneo di test.
    /// </summary>
    public bool CanWriteToFolder(string folderPath)
    {
        _logger.Log($"Test scrittura cartella: {folderPath}");

        try
        {
            // Crea la cartella se non esiste
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
                _logger.Log("Cartella creata");
            }

            // Prova a creare e eliminare un file di test
            var testFile = Path.Combine(folderPath, $".write_test_{Guid.NewGuid():N}.tmp");
            File.WriteAllText(testFile, "test");
            File.Delete(testFile);

            _logger.Log("Test scrittura superato");
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Test scrittura fallito: {ex.Message}");
            return false;
        }
    }

    private bool TryDeleteFile(string path)
    {
        if (!File.Exists(path)) return true;

        try
        {
            File.SetAttributes(path, FileAttributes.Normal);
            File.Delete(path);
            _logger.Log($"File eliminato: {path}");
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Impossibile eliminare file '{path}': {ex.Message}");
            return false;
        }
    }

    private bool TryDeleteDirectory(string path)
    {
        if (!Directory.Exists(path)) return true;

        try
        {
            Directory.Delete(path, recursive: true);
            _logger.Log($"Cartella eliminata: {path}");
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Impossibile eliminare cartella '{path}': {ex.Message}");
            return false;
        }
    }
}
