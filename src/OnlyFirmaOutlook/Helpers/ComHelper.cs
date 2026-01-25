using System.Runtime.InteropServices;
using OnlyFirmaOutlook.Services;

namespace OnlyFirmaOutlook.Helpers;

/// <summary>
/// Utility per la gestione e pulizia di oggetti COM.
/// Garantisce il rilascio corretto delle istanze COM per evitare processi zombie.
/// </summary>
public static class ComHelper
{
    private static readonly LoggingService _logger = LoggingService.Instance;

    /// <summary>
    /// Rilascia un oggetto COM in modo sicuro.
    /// </summary>
    public static void ReleaseComObject(object? comObject)
    {
        if (comObject == null) return;

        try
        {
            Marshal.FinalReleaseComObject(comObject);
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Errore durante il rilascio di un oggetto COM: {ex.Message}");
        }
    }

    /// <summary>
    /// Rilascia multipli oggetti COM.
    /// </summary>
    public static void ReleaseComObjects(params object?[] comObjects)
    {
        foreach (var obj in comObjects)
        {
            ReleaseComObject(obj);
        }
    }

    /// <summary>
    /// Forza la garbage collection per rilasciare definitivamente i COM objects.
    /// Da usare dopo aver chiamato ReleaseComObject.
    /// </summary>
    public static void ForceGarbageCollection()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    /// <summary>
    /// Chiude e rilascia un documento Word.
    /// </summary>
    public static void CloseWordDocument(dynamic? document, bool saveChanges = false)
    {
        if (document == null) return;

        try
        {
            document.Close(SaveChanges: saveChanges);
            _logger.Log($"Documento Word chiuso (SaveChanges: {saveChanges})");
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Errore durante la chiusura del documento Word: {ex.Message}");
        }
        finally
        {
            ReleaseComObject(document);
        }
    }

    /// <summary>
    /// Chiude e rilascia un'applicazione Word.
    /// </summary>
    public static void QuitWordApplication(dynamic? wordApp, bool saveChanges = false)
    {
        if (wordApp == null) return;

        try
        {
            wordApp.Quit(SaveChanges: saveChanges);
            _logger.Log($"Applicazione Word chiusa (SaveChanges: {saveChanges})");
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Errore durante la chiusura di Word: {ex.Message}");
        }
        finally
        {
            ReleaseComObject(wordApp);
        }
    }

    /// <summary>
    /// Esegue cleanup completo: chiude documento, chiude Word, forza GC.
    /// </summary>
    public static void FullWordCleanup(dynamic? document, dynamic? wordApp, bool saveChanges = false)
    {
        CloseWordDocument(document, saveChanges);
        QuitWordApplication(wordApp, saveChanges);
        ForceGarbageCollection();
        _logger.Log("Cleanup COM Word completato");
    }
}
