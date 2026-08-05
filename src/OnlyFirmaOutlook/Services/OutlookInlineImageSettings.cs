using Microsoft.Win32;

namespace OnlyFirmaOutlook.Services;

/// <summary>
/// Configura Outlook Classic affinché trasformi le immagini locali delle firme
/// in allegati inline quando il messaggio viene inviato.
/// </summary>
public static class OutlookInlineImageSettings
{
    private static readonly string[] OfficeVersions = ["16.0", "15.0", "14.0"];
    private const string ValueName = "Send Pictures With Document";

    public static void Enable(LoggingService logger)
    {
        foreach (var officeVersion in OfficeVersions)
        {
            try
            {
                using var mailOptions = Registry.CurrentUser.CreateSubKey(
                    $@"Software\Microsoft\Office\{officeVersion}\Outlook\Options\Mail",
                    writable: true);

                mailOptions?.SetValue(ValueName, 1, RegistryValueKind.DWord);
            }
            catch (Exception ex)
            {
                logger.LogWarning($"Impossibile abilitare l'invio immagini inline per Outlook {officeVersion}: {ex.Message}");
            }
        }

        logger.Log("Configurazione Outlook per immagini firma inline applicata. Riavvia Outlook per renderla effettiva.");
    }
}
