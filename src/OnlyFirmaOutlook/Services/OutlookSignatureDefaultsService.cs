using Microsoft.Win32;

namespace OnlyFirmaOutlook.Services;

public class OutlookSignatureDefaultsService
{
    private static readonly string[] OutlookVersions = { "16.0", "15.0", "14.0" };
    private readonly LoggingService _logger;

    public OutlookSignatureDefaultsService()
    {
        _logger = LoggingService.Instance;
    }

    public bool TrySetDefaultSignatures(string signatureName, bool setNewMessages, bool setReplies, out string message)
    {
        if (string.IsNullOrWhiteSpace(signatureName))
        {
            message = "Nome firma non valido.";
            return false;
        }

        if (!setNewMessages && !setReplies)
        {
            message = "Nessuna opzione selezionata per la firma predefinita.";
            return false;
        }

        if (TrySetDefaultSignaturesViaWord(signatureName, setNewMessages, setReplies, out message))
        {
            return true;
        }

        return TrySetDefaultSignaturesViaRegistry(signatureName, setNewMessages, setReplies, out message);
    }

    public bool TryClearDefaultSignatures(out string message)
    {
        if (TrySetDefaultSignaturesViaWord(string.Empty, true, true, out message))
        {
            return true;
        }

        return TryClearDefaultSignaturesViaRegistry(out message);
    }

    private bool TrySetDefaultSignaturesViaWord(
        string signatureName,
        bool setNewMessages,
        bool setReplies,
        out string message)
    {
        dynamic? wordApp = null;

        try
        {
            var wordType = Type.GetTypeFromProgID("Word.Application");
            if (wordType == null)
            {
                message = "Microsoft Word non è installato o non accessibile.";
                return false;
            }

            wordApp = Activator.CreateInstance(wordType);
            if (wordApp == null)
            {
                message = "Impossibile creare istanza di Word.";
                return false;
            }

            wordApp.Visible = false;

            var emailSignature = wordApp.EmailOptions.EmailSignature;

            if (setNewMessages)
            {
                emailSignature.NewMessageSignature = signatureName;
                _logger.Log($"Firma predefinita per nuovi messaggi impostata (Word): {signatureName}");
            }

            if (setReplies)
            {
                emailSignature.ReplyMessageSignature = signatureName;
                _logger.Log($"Firma predefinita per risposte/inoltri impostata (Word): {signatureName}");
            }

            message = "Impostazioni predefinite aggiornate.";
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError("Errore impostazione firma predefinita via Word", ex);
            message = $"Errore durante l'aggiornamento: {ex.Message}";
            return false;
        }
        finally
        {
            if (wordApp != null)
            {
                try
                {
                    wordApp.Quit(false);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning($"Errore chiusura Word dopo impostazione firma: {ex.Message}");
                }
            }
        }
    }

    private bool TrySetDefaultSignaturesViaRegistry(
        string signatureName,
        bool setNewMessages,
        bool setReplies,
        out string message)
    {
        try
        {
            var version = ResolveOutlookVersion();
            if (version == null)
            {
                message = "Versione Outlook non rilevata.";
                return false;
            }

            using var key = Registry.CurrentUser.CreateSubKey(
                $@"Software\Microsoft\Office\{version}\Common\MailSettings",
                writable: true);

            if (key == null)
            {
                message = "Impossibile accedere alle impostazioni di Outlook.";
                return false;
            }

            if (setNewMessages)
            {
                key.SetValue("NewSignature", signatureName);
                _logger.Log($"Firma predefinita per nuovi messaggi impostata: {signatureName}");
            }

            if (setReplies)
            {
                key.SetValue("ReplySignature", signatureName);
                _logger.Log($"Firma predefinita per risposte/inoltri impostata: {signatureName}");
            }

            message = "Impostazioni predefinite aggiornate.";
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError("Errore impostazione firma predefinita", ex);
            message = $"Errore durante l'aggiornamento: {ex.Message}";
            return false;
        }
    }

    private bool TryClearDefaultSignaturesViaRegistry(out string message)
    {
        try
        {
            var version = ResolveOutlookVersion();
            if (version == null)
            {
                message = "Versione Outlook non rilevata.";
                return false;
            }

            using var key = Registry.CurrentUser.CreateSubKey(
                $@"Software\Microsoft\Office\{version}\Common\MailSettings",
                writable: true);

            if (key == null)
            {
                message = "Impossibile accedere alle impostazioni di Outlook.";
                return false;
            }

            key.DeleteValue("NewSignature", throwOnMissingValue: false);
            key.DeleteValue("ReplySignature", throwOnMissingValue: false);

            _logger.Log("Impostazioni firma predefinita rimosse dal registro.");
            message = "Impostazioni predefinite rimosse.";
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError("Errore rimozione firma predefinita", ex);
            message = $"Errore durante la rimozione: {ex.Message}";
            return false;
        }
    }

    private string? ResolveOutlookVersion()
    {
        foreach (var version in OutlookVersions)
        {
            using var key = Registry.CurrentUser.OpenSubKey(
                $@"Software\Microsoft\Office\{version}\Common\MailSettings",
                writable: false);
            if (key != null)
            {
                return version;
            }
        }

        foreach (var version in OutlookVersions)
        {
            using var key64 = Registry.LocalMachine.OpenSubKey(
                $@"SOFTWARE\Microsoft\Office\{version}\Outlook",
                writable: false);
            if (key64 != null)
            {
                return version;
            }

            using var key32 = Registry.LocalMachine.OpenSubKey(
                $@"SOFTWARE\WOW6432Node\Microsoft\Office\{version}\Outlook",
                writable: false);
            if (key32 != null)
            {
                return version;
            }
        }

        return null;
    }
}
