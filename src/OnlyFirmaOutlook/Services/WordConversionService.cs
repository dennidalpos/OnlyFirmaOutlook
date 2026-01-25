using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

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
        _logger.Log($"Avvio conversione documento: {sourceDocPath}");
        _logger.Log($"Cartella destinazione: {destinationFolder}");
        _logger.Log($"Nome firma: {signatureName}");
        _logger.Log($"Tipo HTML: {(useFilteredHtml ? "Filtrato" : "Completo")}");

        var result = new ConversionResult();
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

            
            var basePath = Path.Combine(destinationFolder, signatureName);
            var htmPath = basePath + ".htm";
            var rtfPath = basePath + ".rtf";
            var txtPath = basePath + ".txt";

            
            _logger.Log("Creazione istanza Word.Application...");
            var wordType = Type.GetTypeFromProgID("Word.Application");
            if (wordType == null)
            {
                result.ErrorMessage = "Microsoft Word non è installato o non accessibile";
                _logger.LogError(result.ErrorMessage);
                return result;
            }

            wordApp = Activator.CreateInstance(wordType);
            if (wordApp == null)
            {
                result.ErrorMessage = "Impossibile creare istanza di Word";
                _logger.LogError(result.ErrorMessage);
                return result;
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
                result.ErrorMessage = "Impossibile aprire il documento Word";
                _logger.LogError(result.ErrorMessage);
                return result;
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

            NormalizeHtmlImageReferences(htmPath, result.AssetsFolderPath);

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

        return result;
    }

    private void NormalizeHtmlImageReferences(string htmPath, string? assetsFolderPath)
    {
        if (string.IsNullOrWhiteSpace(assetsFolderPath) || !File.Exists(htmPath))
        {
            return;
        }

        var assetsFolderName = Path.GetFileName(assetsFolderPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
        if (string.IsNullOrWhiteSpace(assetsFolderName) || !Directory.Exists(assetsFolderPath))
        {
            return;
        }

        var html = File.ReadAllText(htmPath);
        var regex = new Regex("(?<attr>src|href)\\s*=\\s*(?<quote>[\"'])(?<value>[^\"']+)(\\k<quote>)", RegexOptions.IgnoreCase);
        var updated = false;

        var updatedHtml = regex.Replace(html, match =>
        {
            var value = match.Groups["value"].Value;
            var normalized = NormalizeAssetReference(value, assetsFolderPath, assetsFolderName);

            if (normalized == null || normalized == value)
            {
                return match.Value;
            }

            updated = true;
            return $"{match.Groups["attr"].Value}={match.Groups["quote"].Value}{normalized}{match.Groups["quote"].Value}";
        });

        if (updated)
        {
            File.WriteAllText(htmPath, updatedHtml);
            _logger.Log("Riferimenti immagini HTML normalizzati per embed Outlook");
        }
    }

    private static string? NormalizeAssetReference(string value, string assetsFolderPath, string assetsFolderName)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        if (value.StartsWith("cid:", StringComparison.OrdinalIgnoreCase) ||
            value.StartsWith("data:", StringComparison.OrdinalIgnoreCase) ||
            value.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
            value.StartsWith("https://", StringComparison.OrdinalIgnoreCase) ||
            value.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase))
        {
            return value;
        }

        var localPath = value;
        if (value.StartsWith("file:", StringComparison.OrdinalIgnoreCase) && Uri.TryCreate(value, UriKind.Absolute, out var uri) && uri.IsFile)
        {
            localPath = uri.LocalPath;
        }

        var fileName = Path.GetFileName(localPath);
        if (string.IsNullOrWhiteSpace(fileName))
        {
            return value;
        }

        if (value.Contains(assetsFolderName, StringComparison.OrdinalIgnoreCase))
        {
            return $"{assetsFolderName}/{fileName}";
        }

        if (Path.IsPathRooted(localPath))
        {
            var assetsFilePath = Path.Combine(assetsFolderPath, fileName);
            if (File.Exists(assetsFilePath))
            {
                return $"{assetsFolderName}/{fileName}";
            }
        }
        else
        {
            var assetsFilePath = Path.Combine(assetsFolderPath, fileName);
            if (File.Exists(assetsFilePath))
            {
                return $"{assetsFolderName}/{fileName}";
            }
        }

        return value;
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
        GC.Collect();
        GC.WaitForPendingFinalizers();

        _logger.Log("Cleanup COM completato");
    }

    
    
    
    
    public static string SanitizeFileName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            return "Firma";

        var invalidChars = Path.GetInvalidFileNameChars();
        var sanitized = new string(name
            .Select(c => invalidChars.Contains(c) ? '_' : c)
            .ToArray());

        
        while (sanitized.Contains("__"))
        {
            sanitized = sanitized.Replace("__", "_");
        }

        
        sanitized = sanitized.Trim('_', ' ');

        
        if (string.IsNullOrWhiteSpace(sanitized))
            return "Firma";

        
        if (sanitized.Length > 100)
            sanitized = sanitized[..100];

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
}
