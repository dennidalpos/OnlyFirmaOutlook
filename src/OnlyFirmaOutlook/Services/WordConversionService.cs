// (c) 2026 Danny Perondi. All rights reserved. Proprietary and confidential.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;

namespace OnlyFirmaOutlook.Services;

public class WordConversionService
{
    private readonly LoggingService _logger;

    
    private const int WdFormatFilteredHTML = 10;
    private const int WdFormatHTML = 8;
    private const int WdFormatRTF = 6;
    private const int WdFormatText = 2;

    public WordConversionService()
    {
        _logger = LoggingService.Instance;
    }

    
    
    
    public class ConversionResult
    {
        public bool Success { get; set; }
        public string? HtmFilePath { get; set; }
        public string? RtfFilePath { get; set; }
        public string? TxtFilePath { get; set; }
        public string? AssetsFolderPath { get; set; }
        public string? ErrorMessage { get; set; }
    }

    
    
    
    
    
    
    
    
    public ConversionResult ConvertDocument(
        string sourceDocPath,
        string destinationFolder,
        string signatureName,
        bool useFilteredHtml = true)
    {
        var sanitizedSignatureName = SanitizeFileName(signatureName);
        if (!string.Equals(signatureName, sanitizedSignatureName, StringComparison.Ordinal))
        {
            _logger.LogWarning($"Nome firma normalizzato per export: '{signatureName}' → '{sanitizedSignatureName}'");
            signatureName = sanitizedSignatureName;
        }

        _logger.Log($"Avvio conversione documento: {sourceDocPath}");
        _logger.Log($"Cartella destinazione: {destinationFolder}");
        _logger.Log($"Nome firma: {signatureName}");
        _logger.Log($"Tipo HTML: {(useFilteredHtml ? "Filtrato" : "Completo")}");
        var result = new ConversionResult();
        string? stagingFolder = null;
        dynamic? wordApp = null;
        dynamic? doc = null;

        try
        {
            
            if (!File.Exists(sourceDocPath))
            {
                result.ErrorMessage = $"File sorgente non trovato: {sourceDocPath}";
                _logger.LogError(result.ErrorMessage);
                return result;
            }

            
            if (!Directory.Exists(destinationFolder))
            {
                Directory.CreateDirectory(destinationFolder);
                _logger.Log("Cartella destinazione creata");
            }

            stagingFolder = CreateStagingFolder();
            var basePath = Path.Combine(stagingFolder, signatureName);
            var htmPath = basePath + ".htm";
            var rtfPath = basePath + ".rtf";
            var txtPath = basePath + ".txt";

            
            _logger.Log("Creazione istanza Word.Application...");
            var wordType = Type.GetTypeFromProgID("Word.Application");
            if (wordType == null)
            {
                throw new InvalidOperationException("Microsoft Word non è installato o non accessibile");
            }

            wordApp = Activator.CreateInstance(wordType);
            if (wordApp == null)
            {
                throw new InvalidOperationException("Impossibile creare istanza di Word");
            }

            wordApp.Visible = false;
            wordApp.DisplayAlerts = 0; 

            
            _logger.Log("Apertura documento...");
            doc = wordApp.Documents.Open(
                FileName: sourceDocPath,
                ReadOnly: true,
                AddToRecentFiles: false,
                Visible: false);

            if (doc == null)
            {
                throw new InvalidOperationException("Impossibile aprire il documento Word");
            }

            _logger.Log("Documento aperto con successo");

            
            _logger.Log($"Salvataggio HTML ({(useFilteredHtml ? "filtrato" : "completo")})...");
            var htmlFormat = useFilteredHtml ? WdFormatFilteredHTML : WdFormatHTML;
            doc.SaveAs2(
                FileName: htmPath,
                FileFormat: htmlFormat,
                AddToRecentFiles: false);
            result.HtmFilePath = htmPath;
            _logger.Log($"HTML salvato: {htmPath}");

            
            _logger.Log("Salvataggio RTF...");
            doc.SaveAs2(
                FileName: rtfPath,
                FileFormat: WdFormatRTF,
                AddToRecentFiles: false);
            result.RtfFilePath = rtfPath;
            _logger.Log($"RTF salvato: {rtfPath}");

            
            _logger.Log("Salvataggio TXT...");
            doc.SaveAs2(
                FileName: txtPath,
                FileFormat: WdFormatText,
                AddToRecentFiles: false);
            result.TxtFilePath = txtPath;
            _logger.Log($"TXT salvato: {txtPath}");

            
            var filesFolderPath = basePath + "_files";
            var fileFolderPath = basePath + "_file";

            if (Directory.Exists(filesFolderPath))
            {
                result.AssetsFolderPath = filesFolderPath;
                _logger.Log($"Cartella assets trovata: {filesFolderPath}");
            }
            else if (Directory.Exists(fileFolderPath))
            {
                result.AssetsFolderPath = fileFolderPath;
                _logger.Log($"Cartella assets trovata: {fileFolderPath}");
            }
            else
            {
                _logger.Log("Nessuna cartella assets generata (il documento potrebbe non contenere immagini)");
            }

            result.Success = true;
            _logger.Log("Conversione completata con successo");
        }
        catch (COMException comEx)
        {
            result.ErrorMessage = $"Errore COM durante la conversione: {comEx.Message} (0x{comEx.ErrorCode:X8})";
            _logger.LogError(result.ErrorMessage, comEx);

            
            if (comEx.ErrorCode == unchecked((int)0x800A175D))
            {
                result.ErrorMessage += "\n\nIl file potrebbe essere in 'Protected View'. " +
                    "Aprire il file manualmente in Word, abilitare la modifica e riprovare.";
            }
        }
        catch (Exception ex)
        {
            result.ErrorMessage = $"Errore durante la conversione: {ex.Message}";
            _logger.LogError(result.ErrorMessage, ex);
        }
        finally
        {
            
            CleanupComObjects(doc, wordApp);
        }

        if (result.Success && result.HtmFilePath != null && stagingFolder != null)
        {
            if (!TryFinalizeConversion(result, stagingFolder, signatureName) ||
                !TryCommitConversion(result, stagingFolder, destinationFolder, signatureName))
            {
                result.Success = false;
            }
        }

        if (stagingFolder != null)
        {
            CleanupDirectory(stagingFolder, "cartella di staging conversione");
        }

        return result;
    }

    protected virtual string? ReadHtmlForPostProcessing(string path)
    {
        return ReadAllTextWithRetry(path);
    }

    protected virtual string InlineCss(string html)
    {
        var cssInliner = new CssInliner();
        return cssInliner.InlineCss(html);
    }

    protected virtual string NormalizeHtml(string html)
    {
        var normalizer = new WordHtmlSignatureNormalizer();
        return normalizer.Normalize(html);
    }

    protected virtual AssetProcessingResult ProcessAssets(string html, string sourceHtmlPath, string assetsFolderPath)
    {
        var assetManager = new AssetManager();
        return assetManager.ProcessImages(html, sourceHtmlPath, assetsFolderPath);
    }

    protected virtual void InstallSignature(string destinationFolder, string signatureName, string html, string plainText)
    {
        var installer = new SignatureInstaller();
        installer.Install(destinationFolder, signatureName, html, plainText);
    }

    private void CleanupWordAssetFolders(string destinationFolder, string signatureName)
    {
        var basePath = Path.Combine(destinationFolder, signatureName);
        var filesFolderPath = basePath + "_files";
        var fileFolderPath = basePath + "_file";

        if (Directory.Exists(filesFolderPath) && Directory.Exists(fileFolderPath))
        {
            try
            {
                Directory.Delete(fileFolderPath, true);
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Impossibile eliminare cartella assets duplicata: {ex.Message}");
            }
        }
    }

    private static string? ReadAllTextWithRetry(string path)
    {
        const int maxAttempts = 5;
        const int delayMs = 150;

        for (var attempt = 0; attempt < maxAttempts; attempt++)
        {
            try
            {
                using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var reader = new StreamReader(stream);
                return reader.ReadToEnd();
            }
            catch (IOException)
            {
                Thread.Sleep(delayMs);
            }
        }

        return null;
    }

    private void CleanupComObjects(dynamic? doc, dynamic? wordApp)
    {
        _logger.Log("Cleanup oggetti COM...");

        try
        {
            if (doc != null)
            {
                try
                {
                    doc.Close(SaveChanges: false);
                    _logger.Log("Documento chiuso");
                }
                catch (Exception ex)
                {
                    _logger.LogWarning($"Errore chiusura documento: {ex.Message}");
                }
                finally
                {
                    Marshal.FinalReleaseComObject(doc);
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Errore rilascio documento COM: {ex.Message}");
        }

        try
        {
            if (wordApp != null)
            {
                try
                {
                    wordApp.Quit(SaveChanges: false);
                    _logger.Log("Word chiuso");
                }
                catch (Exception ex)
                {
                    _logger.LogWarning($"Errore chiusura Word: {ex.Message}");
                }
                finally
                {
                    Marshal.FinalReleaseComObject(wordApp);
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Errore rilascio Word COM: {ex.Message}");
        }

        GC.Collect();
        GC.WaitForPendingFinalizers();

        _logger.Log("Cleanup COM completato");
    }

    
    
    
    
    private static readonly HashSet<char> InvalidFileNameChars = new(
        Path.GetInvalidFileNameChars().Concat(new[] { '<', '>', '"', ':', '/', '\\', '|', '?', '*' }));

    public static string SanitizeFileName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return "Firma";
        }

        var sanitized = new string(name
            .Select(c => InvalidFileNameChars.Contains(c) ? '_' : c)
            .ToArray());

        
        while (sanitized.Contains("__"))
        {
            sanitized = sanitized.Replace("__", "_");
        }

        
        sanitized = sanitized.Trim('_', ' ');

        
        if (string.IsNullOrWhiteSpace(sanitized))
        {
            return "Firma";
        }

        
        if (sanitized.Length > 100)
        {
            sanitized = sanitized[..100];
        }

        return sanitized;
    }

    
    
    
    public static string GenerateSignatureName(string baseName, string? identifier)
    {
        var sanitizedBase = SanitizeFileName(baseName);

        if (string.IsNullOrWhiteSpace(identifier))
        {
            return sanitizedBase;
        }

        var sanitizedIdentifier = SanitizeFileName(identifier);
        return $"{sanitizedBase} ({sanitizedIdentifier})";
    }

    private bool TryFinalizeConversion(
        ConversionResult result,
        string destinationFolder,
        string signatureName)
    {
        try
        {
            var html = ReadHtmlForPostProcessing(result.HtmFilePath!);
            if (html == null)
            {
                result.Success = false;
                result.ErrorMessage = "Impossibile leggere HTML firma per normalizzazione.";
                _logger.LogWarning(result.ErrorMessage);
                return false;
            }

            var inlined = InlineCss(html);
            var normalized = NormalizeHtml(inlined);

            var assetsFolder = Path.Combine(destinationFolder, $"{signatureName}_files");
            var assetResult = ProcessAssets(normalized, result.HtmFilePath!, assetsFolder);
            InstallSignature(destinationFolder, signatureName, assetResult.Html, assetResult.PlainText);

            if (Directory.Exists(assetsFolder) && !Directory.EnumerateFileSystemEntries(assetsFolder).Any())
            {
                try
                {
                    Directory.Delete(assetsFolder, true);
                }
                catch
                {
                }

                result.AssetsFolderPath = null;
            }
            else
            {
                result.AssetsFolderPath = assetsFolder;
            }

            result.HtmFilePath = Path.Combine(destinationFolder, signatureName + ".htm");
            result.TxtFilePath = Path.Combine(destinationFolder, signatureName + ".txt");

            CleanupWordAssetFolders(destinationFolder, signatureName);
            return true;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Errore normalizzazione HTML firma: {ex.Message}";
            _logger.LogWarning(result.ErrorMessage);
            return false;
        }
    }

    private string CreateStagingFolder()
    {
        var stagingFolder = Path.Combine(
            Path.GetTempPath(),
            "OnlyFirmaOutlook",
            "Conversion",
            Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(stagingFolder);
        _logger.Log($"Cartella di staging creata: {stagingFolder}");
        return stagingFolder;
    }

    private bool TryCommitConversion(
        ConversionResult result,
        string stagingFolder,
        string destinationFolder,
        string signatureName)
    {
        var artifactNames = GetArtifactNames(signatureName);
        var rollbackFolder = Path.Combine(
            Path.GetTempPath(),
            "OnlyFirmaOutlook",
            "ConversionRollback",
            Guid.NewGuid().ToString("N"));

        try
        {
            Directory.CreateDirectory(rollbackFolder);

            foreach (var artifactName in artifactNames)
            {
                MoveArtifactIfPresent(
                    Path.Combine(destinationFolder, artifactName),
                    Path.Combine(rollbackFolder, artifactName));
            }

            foreach (var artifactName in artifactNames)
            {
                MoveArtifactIfPresent(
                    Path.Combine(stagingFolder, artifactName),
                    Path.Combine(destinationFolder, artifactName));
            }

            result.HtmFilePath = Path.Combine(destinationFolder, signatureName + ".htm");
            result.RtfFilePath = Path.Combine(destinationFolder, signatureName + ".rtf");
            result.TxtFilePath = Path.Combine(destinationFolder, signatureName + ".txt");
            result.AssetsFolderPath = Directory.Exists(Path.Combine(destinationFolder, signatureName + "_files"))
                ? Path.Combine(destinationFolder, signatureName + "_files")
                : null;

            CleanupDirectory(rollbackFolder, "backup rollback conversione");
            _logger.Log("Conversione pubblicata nella cartella destinazione");
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError("Pubblicazione conversione fallita: ripristino della firma precedente", ex);

            RestorePreviousArtifacts(artifactNames, rollbackFolder, destinationFolder);

            result.ErrorMessage = $"Impossibile pubblicare la firma: {ex.Message}";
            return false;
        }
        finally
        {
            CleanupDirectory(rollbackFolder, "backup rollback conversione");
        }
    }

    private void RestorePreviousArtifacts(
        IEnumerable<string> artifactNames,
        string rollbackFolder,
        string destinationFolder)
    {
        foreach (var artifactName in artifactNames)
        {
            try
            {
                DeleteArtifactIfPresent(Path.Combine(destinationFolder, artifactName));
                MoveArtifactIfPresent(
                    Path.Combine(rollbackFolder, artifactName),
                    Path.Combine(destinationFolder, artifactName));
            }
            catch (Exception ex)
            {
                _logger.LogError($"Rollback artefatto fallito: {artifactName}", ex);
            }
        }
    }

    private static string[] GetArtifactNames(string signatureName) =>
    [
        signatureName + ".htm",
        signatureName + ".rtf",
        signatureName + ".txt",
        signatureName + "_files",
        signatureName + "_file"
    ];

    private static void MoveArtifactIfPresent(string sourcePath, string destinationPath)
    {
        if (Directory.Exists(sourcePath))
        {
            Directory.Move(sourcePath, destinationPath);
        }
        else if (File.Exists(sourcePath))
        {
            File.Move(sourcePath, destinationPath);
        }
    }

    private static void DeleteArtifactIfPresent(string path)
    {
        if (Directory.Exists(path))
        {
            Directory.Delete(path, recursive: true);
        }
        else if (File.Exists(path))
        {
            File.SetAttributes(path, FileAttributes.Normal);
            File.Delete(path);
        }
    }

    private void CleanupDirectory(string path, string description)
    {
        try
        {
            if (Directory.Exists(path))
            {
                Directory.Delete(path, recursive: true);
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Impossibile eliminare {description}: {ex.Message}");
        }
    }

}
