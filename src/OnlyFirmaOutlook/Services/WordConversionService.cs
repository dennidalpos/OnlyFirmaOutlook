using System.IO;
using System.Runtime.InteropServices;

namespace OnlyFirmaOutlook.Services;

/// <summary>
/// Servizio per la conversione di documenti Word in firme (HTML, RTF, TXT).
/// Utilizza COM Interop con Microsoft.Office.Interop.Word.
/// Tutte le operazioni COM devono essere eseguite su thread STA.
/// </summary>
public class WordConversionService
{
    private readonly LoggingService _logger;
    private readonly SignatureRepository _signatureRepository;

    // Costanti per i formati di salvataggio Word
    private const int WdFormatFilteredHTML = 10;
    private const int WdFormatHTML = 8;
    private const int WdFormatRTF = 6;
    private const int WdFormatText = 2;

    public WordConversionService()
    {
        _logger = LoggingService.Instance;
        _signatureRepository = new SignatureRepository();
    }

    /// <summary>
    /// Risultato della conversione.
    /// </summary>
    public class ConversionResult
    {
        public bool Success { get; set; }
        public string? HtmFilePath { get; set; }
        public string? RtfFilePath { get; set; }
        public string? TxtFilePath { get; set; }
        public string? AssetsFolderPath { get; set; }
        public string? ErrorMessage { get; set; }
    }

    /// <summary>
    /// Converte un documento Word in firma Outlook.
    /// </summary>
    /// <param name="sourceDocPath">Percorso del documento Word sorgente</param>
    /// <param name="destinationFolder">Cartella di destinazione</param>
    /// <param name="signatureName">Nome della firma (sanitizzato)</param>
    /// <param name="useFilteredHtml">true per HTML filtrato, false per HTML completo</param>
    /// <returns>Risultato della conversione</returns>
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
            // Verifica che il file sorgente esista
            if (!File.Exists(sourceDocPath))
            {
                result.ErrorMessage = $"File sorgente non trovato: {sourceDocPath}";
                _logger.LogError(result.ErrorMessage);
                return result;
            }

            // Crea la cartella di destinazione se non esiste
            if (!Directory.Exists(destinationFolder))
            {
                Directory.CreateDirectory(destinationFolder);
                _logger.Log("Cartella destinazione creata");
            }

            // Calcola i percorsi di output
            var basePath = Path.Combine(destinationFolder, signatureName);
            var htmPath = basePath + ".htm";
            var rtfPath = basePath + ".rtf";
            var txtPath = basePath + ".txt";

            // Crea l'istanza di Word
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
            wordApp.DisplayAlerts = 0; // wdAlertsNone

            // Apri il documento in sola lettura
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

            // Salva come HTML
            _logger.Log($"Salvataggio HTML ({(useFilteredHtml ? "filtrato" : "completo")})...");
            var htmlFormat = useFilteredHtml ? WdFormatFilteredHTML : WdFormatHTML;
            doc.SaveAs2(
                FileName: htmPath,
                FileFormat: htmlFormat,
                AddToRecentFiles: false);
            result.HtmFilePath = htmPath;
            _logger.Log($"HTML salvato: {htmPath}");

            // Salva come RTF
            _logger.Log("Salvataggio RTF...");
            doc.SaveAs2(
                FileName: rtfPath,
                FileFormat: WdFormatRTF,
                AddToRecentFiles: false);
            result.RtfFilePath = rtfPath;
            _logger.Log($"RTF salvato: {rtfPath}");

            // Salva come TXT
            _logger.Log("Salvataggio TXT...");
            doc.SaveAs2(
                FileName: txtPath,
                FileFormat: WdFormatText,
                AddToRecentFiles: false);
            result.TxtFilePath = txtPath;
            _logger.Log($"TXT salvato: {txtPath}");

            // Cerca la cartella assets generata da Word
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

            // Gestione errori COM specifici
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
            // Cleanup COM objects
            CleanupComObjects(doc, wordApp);
        }

        return result;
    }

    /// <summary>
    /// Pulisce gli oggetti COM per evitare processi zombie.
    /// </summary>
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

        // Forza garbage collection per rilasciare definitivamente i COM objects
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        GC.WaitForPendingFinalizers();

        _logger.Log("Cleanup COM completato");
    }

    /// <summary>
    /// Sanitizza un nome per l'uso come nome file/cartella.
    /// Rimuove o sostituisce i caratteri non validi.
    /// </summary>
    public static string SanitizeFileName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            return "Firma";

        var invalidChars = Path.GetInvalidFileNameChars();
        var sanitized = new string(name
            .Select(c => invalidChars.Contains(c) ? '_' : c)
            .ToArray());

        // Rimuovi underscore multipli consecutivi
        while (sanitized.Contains("__"))
        {
            sanitized = sanitized.Replace("__", "_");
        }

        // Rimuovi underscore iniziali e finali
        sanitized = sanitized.Trim('_', ' ');

        // Se il nome è vuoto dopo la sanitizzazione, usa un default
        if (string.IsNullOrWhiteSpace(sanitized))
            return "Firma";

        // Limita la lunghezza
        if (sanitized.Length > 100)
            sanitized = sanitized[..100];

        return sanitized;
    }

    /// <summary>
    /// Genera il nome completo della firma includendo l'identificativo.
    /// </summary>
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
