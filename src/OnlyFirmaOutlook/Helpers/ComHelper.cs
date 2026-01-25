using System.Runtime.InteropServices;
using OnlyFirmaOutlook.Services;

namespace OnlyFirmaOutlook.Helpers;





public static class ComHelper
{
    private static readonly LoggingService _logger = LoggingService.Instance;

    
    
    
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

    
    
    
    public static void ReleaseComObjects(params object?[] comObjects)
    {
        foreach (var obj in comObjects)
        {
            ReleaseComObject(obj);
        }
    }

    
    
    
    
    public static void ForceGarbageCollection()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    
    
    
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

    
    
    
    public static void FullWordCleanup(dynamic? document, dynamic? wordApp, bool saveChanges = false)
    {
        CloseWordDocument(document, saveChanges);
        QuitWordApplication(wordApp, saveChanges);
        ForceGarbageCollection();
        _logger.Log("Cleanup COM Word completato");
    }
}
