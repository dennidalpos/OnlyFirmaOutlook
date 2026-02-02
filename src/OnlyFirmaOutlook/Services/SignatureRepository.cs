using System;
using System.IO;
using System.IO.Compression;
using OnlyFirmaOutlook.Models;

namespace OnlyFirmaOutlook.Services;





public class SignatureRepository
{
    private readonly LoggingService _logger;

    public SignatureRepository()
    {
        _logger = LoggingService.Instance;
    }

    
    
    
    public static string GetDefaultOutlookSignaturesFolder()
    {
        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Microsoft",
            "Signatures");
    }

    
    
    
    public static string GetAlternativeOutputFolder()
    {
        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "OnlyFirmaOutlook",
            "Output");
    }

    
    
    
    
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
            var signatureNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var htmFiles = Directory.GetFiles(folderPath, "*.htm", SearchOption.TopDirectoryOnly);
            var rtfFiles = Directory.GetFiles(folderPath, "*.rtf", SearchOption.TopDirectoryOnly);
            var txtFiles = Directory.GetFiles(folderPath, "*.txt", SearchOption.TopDirectoryOnly);

            foreach (var file in htmFiles)
            {
                signatureNames.Add(Path.GetFileNameWithoutExtension(file));
            }

            foreach (var file in rtfFiles)
            {
                signatureNames.Add(Path.GetFileNameWithoutExtension(file));
            }

            foreach (var file in txtFiles)
            {
                signatureNames.Add(Path.GetFileNameWithoutExtension(file));
            }

            foreach (var baseName in signatureNames)
            {
                var signature = new SignatureInfo
                {
                    Name = baseName,
                    FolderPath = folderPath,
                    HasHtm = File.Exists(Path.Combine(folderPath, baseName + ".htm")),
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

        
        success &= TryDeleteFile(basePath + ".htm");

        
        success &= TryDeleteFile(basePath + ".rtf");

        
        success &= TryDeleteFile(basePath + ".txt");

        
        success &= TryDeleteDirectory(basePath + "_files");

        
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

    
    
    
    public bool SignatureExists(string folderPath, string signatureName)
    {
        var htmPath = Path.Combine(folderPath, signatureName + ".htm");
        return File.Exists(htmPath);
    }

    
    
    
    
    public bool CanWriteToFolder(string folderPath)
    {
        _logger.Log($"Test scrittura cartella: {folderPath}");

        try
        {
            
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
                _logger.Log("Cartella creata");
            }

            
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

    public bool CreateBackupInSignaturesFolder()
    {
        var signaturesFolder = GetDefaultOutlookSignaturesFolder();
        if (!Directory.Exists(signaturesFolder))
        {
            _logger.LogWarning($"Cartella firme non trovata per backup: {signaturesFolder}");
            return false;
        }

        var backupPrefix = "backup_firme_onlyfirmaoutlook_";
        var signatureFiles = Directory.EnumerateFiles(signaturesFolder, "*", SearchOption.AllDirectories)
            .Where(file => !Path.GetFileName(file).StartsWith(backupPrefix, StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (signatureFiles.Count == 0)
        {
            _logger.Log("Nessuna firma da salvare: backup non creato.");
            return false;
        }

        var timestamp = DateTime.Now.ToString("yyyy-MM-dd-HH-mm");
        var backupFileName = $"{backupPrefix}{timestamp}.zip";
        var backupPath = Path.Combine(signaturesFolder, backupFileName);

        try
        {
            if (File.Exists(backupPath))
            {
                File.Delete(backupPath);
            }

            using var archive = ZipFile.Open(backupPath, ZipArchiveMode.Create);
            foreach (var file in signatureFiles)
            {
                var relativePath = Path.GetRelativePath(signaturesFolder, file);
                if (string.Equals(relativePath, backupFileName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                archive.CreateEntryFromFile(file, relativePath);
            }

            _logger.Log($"Backup firme creato: {backupPath}");
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Backup firme fallito: {ex.Message}");
            return false;
        }
    }

    public List<BackupInfo> GetBackups(string folderPath)
    {
        var backups = new List<BackupInfo>();

        if (!Directory.Exists(folderPath))
        {
            _logger.LogWarning($"Cartella firme non trovata per lettura backup: {folderPath}");
            return backups;
        }

        try
        {
            var backupFiles = Directory.EnumerateFiles(folderPath, "backup_firme_onlyfirmaoutlook_*.zip", SearchOption.TopDirectoryOnly);

            foreach (var backupFile in backupFiles)
            {
                var info = new FileInfo(backupFile);
                backups.Add(new BackupInfo
                {
                    FileName = info.Name,
                    FullPath = info.FullName,
                    CreatedAt = info.LastWriteTime,
                    SizeBytes = info.Length
                });
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Errore durante lettura backup firme: {ex.Message}");
        }

        return backups
            .OrderByDescending(b => b.CreatedAt)
            .ThenBy(b => b.FileName)
            .ToList();
    }

    public bool DeleteBackup(BackupInfo backup)
    {
        if (backup == null)
        {
            _logger.LogWarning("Tentativo di eliminare backup null");
            return false;
        }

        try
        {
            if (File.Exists(backup.FullPath))
            {
                File.Delete(backup.FullPath);
                _logger.Log($"Backup eliminato: {backup.FullPath}");
            }
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Errore durante eliminazione backup: {ex.Message}");
            return false;
        }
    }

    public bool RestoreBackup(BackupInfo backup, string destinationFolder)
    {
        if (backup == null)
        {
            _logger.LogWarning("Tentativo di ripristinare backup null");
            return false;
        }

        if (!File.Exists(backup.FullPath))
        {
            _logger.LogWarning($"Backup non trovato: {backup.FullPath}");
            return false;
        }

        try
        {
            if (!Directory.Exists(destinationFolder))
            {
                Directory.CreateDirectory(destinationFolder);
            }

            ZipFile.ExtractToDirectory(backup.FullPath, destinationFolder, overwriteFiles: true);
            _logger.Log($"Backup ripristinato: {backup.FullPath}");
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Errore durante ripristino backup: {ex.Message}");
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
