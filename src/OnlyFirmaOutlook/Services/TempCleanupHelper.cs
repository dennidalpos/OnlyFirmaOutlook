using System.IO;
using System.Linq;
using System.Threading;

namespace OnlyFirmaOutlook.Services;

public static class TempCleanupHelper
{
    public static void CleanupDirectoryWithRetries(
        string folderPath,
        LoggingService logger,
        string contextLabel,
        int maxRetries = 3,
        int retryDelayMs = 100)
    {
        if (!Directory.Exists(folderPath))
        {
            logger.Log($"{contextLabel} già eliminata o non esistente: {folderPath}");
            return;
        }

        logger.Log($"Avvio cleanup {contextLabel}: {folderPath}");

        for (var attempt = 1; attempt <= maxRetries; attempt++)
        {
            try
            {
                DeleteContents(folderPath, logger);
                Directory.Delete(folderPath, recursive: true);
                logger.Log($"Cleanup {contextLabel} completato con successo");
                return;
            }
            catch (Exception ex)
            {
                logger.LogWarning($"Tentativo {attempt}/{maxRetries} cleanup {contextLabel} fallito: {ex.Message}");

                if (attempt < maxRetries)
                {
                    Thread.Sleep(retryDelayMs * attempt);
                }
            }
        }

        logger.LogWarning($"Cleanup {contextLabel} non completato dopo {maxRetries} tentativi. " +
                          $"La cartella '{folderPath}' potrebbe richiedere pulizia manuale.");
    }

    private static void DeleteContents(string folderPath, LoggingService logger)
    {
        foreach (var file in Directory.GetFiles(folderPath, "*", SearchOption.AllDirectories))
        {
            try
            {
                File.SetAttributes(file, FileAttributes.Normal);
                File.Delete(file);
            }
            catch (Exception ex)
            {
                logger.LogWarning($"Impossibile eliminare file '{file}': {ex.Message}");
            }
        }

        foreach (var dir in Directory.GetDirectories(folderPath, "*", SearchOption.AllDirectories)
                                     .OrderByDescending(d => d.Length))
        {
            try
            {
                Directory.Delete(dir, recursive: false);
            }
            catch (Exception ex)
            {
                logger.LogWarning($"Impossibile eliminare cartella '{dir}': {ex.Message}");
            }
        }
    }
}
