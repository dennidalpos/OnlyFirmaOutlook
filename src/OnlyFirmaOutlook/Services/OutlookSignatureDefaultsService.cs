using System.Linq;
using System.Text;
using Microsoft.Win32;
using OnlyFirmaOutlook.Models;

namespace OnlyFirmaOutlook.Services;

public class OutlookSignatureDefaultsService
{
    private static readonly string[] OutlookVersions = { "16.0", "15.0", "14.0" };
    private readonly LoggingService _logger;

    public OutlookSignatureDefaultsService()
    {
        _logger = LoggingService.Instance;
    }

    public bool TrySetDefaultSignatures(OutlookAccount account, string signatureName, bool setNewMessages, bool setReplies, out string message)
    {
        if (account == null)
        {
            message = "Account Outlook non valido.";
            return false;
        }

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

        return TrySetDefaultSignaturesViaRegistry(account, signatureName, setNewMessages, setReplies, out message);
    }

    public bool TryClearDefaultSignatures(OutlookAccount account, out string message)
    {
        if (account == null)
        {
            message = "Account Outlook non valido.";
            return false;
        }

        return TryClearDefaultSignaturesViaRegistry(account, out message);
    }

    public bool TryGetDefaultSignatures(OutlookAccount account, out string? newSignature, out string? replySignature, out string message)
    {
        newSignature = null;
        replySignature = null;

        if (account == null)
        {
            message = "Account Outlook non valido.";
            return false;
        }

        try
        {
            var version = ResolveOutlookVersion();
            if (version == null)
            {
                message = "Versione Outlook non rilevata.";
                return false;
            }

            using var key = Registry.CurrentUser.OpenSubKey(
                $@"Software\Microsoft\Office\{version}\Common\MailSettings",
                writable: false);
            if (key == null)
            {
                message = "Impostazioni Outlook non disponibili.";
                return false;
            }

            newSignature = key.GetValue("NewSignature") as string;
            replySignature = key.GetValue("ReplySignature") as string;
            message = "OK";
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError("Errore lettura impostazioni firma predefinita", ex);
            message = $"Errore durante la lettura: {ex.Message}";
            return false;
        }
    }

    public bool CanWriteDefaultSignatureRegistry(OutlookAccount account, out string message)
    {
        if (account == null)
        {
            message = "Account Outlook non valido.";
            return false;
        }

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
                message = "Accesso al registro negato.";
                return false;
            }

            message = "Accesso al registro disponibile.";
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError("Errore accesso registro firma predefinita", ex);
            message = $"Accesso al registro negato: {ex.Message}";
            return false;
        }
    }

    private bool TrySetDefaultSignaturesViaRegistry(
        OutlookAccount account,
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

    private bool TryClearDefaultSignaturesViaRegistry(OutlookAccount account, out string message)
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

    private bool TryGetAccountRegistryKeyPath(string version, OutlookAccount account, out string accountKeyPath, out string message)
    {
        accountKeyPath = string.Empty;

        var profileName = ResolveDefaultProfileName(version);
        if (string.IsNullOrWhiteSpace(profileName))
        {
            message = "Profilo Outlook non rilevato.";
            return false;
        }

        var profileRootPath = $@"Software\Microsoft\Office\{version}\Outlook\Profiles\{profileName}";
        var accountsRootPath = $@"{profileRootPath}\9375CFF0413111d3B88A00104B2A6676";

        var identifiers = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            account.SmtpAddress,
            account.OwnerSmtpAddress,
            account.DisplayName
        }.Where(value => !string.IsNullOrWhiteSpace(value)).ToList();

        if (identifiers.Count == 0)
        {
            message = "Identificativo account non disponibile per la ricerca nel registro.";
            return false;
        }

        string? matchPath = null;

        matchPath = FindAccountKeyPath(accountsRootPath, identifiers);

        if (matchPath == null)
        {
            matchPath = FindAccountKeyPath(profileRootPath, identifiers);
        }

        if (matchPath == null)
        {
            message = "Impossibile associare l'account selezionato al profilo Outlook.";
            return false;
        }

        accountKeyPath = matchPath;
        message = "OK";
        return true;
    }

    private static string? FindAccountKeyPath(string rootPath, IReadOnlyCollection<string> identifiers)
    {
        var queue = new Queue<string>();
        queue.Enqueue(rootPath);

        while (queue.Count > 0)
        {
            var currentPath = queue.Dequeue();
            using var currentKey = Registry.CurrentUser.OpenSubKey(currentPath, writable: false);
            if (currentKey == null)
            {
                continue;
            }

            if (IsAccountMatch(currentKey, identifiers))
            {
                return currentPath;
            }

            foreach (var subKeyName in currentKey.GetSubKeyNames())
            {
                var subKeyPath = $@"{currentPath}\{subKeyName}";
                queue.Enqueue(subKeyPath);
            }
        }

        return null;
    }

    private static bool IsAccountMatch(RegistryKey key, IReadOnlyCollection<string> identifiers)
    {
        var candidateValueNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "001f3001",
            "001f39fe",
            "001e3001",
            "Account Name",
            "SMTP Address",
            "Display Name",
            "Email Address"
        };

        var valueNames = key.GetValueNames();
        foreach (var valueName in valueNames)
        {
            if (!candidateValueNames.Contains(valueName))
            {
                continue;
            }

            var value = ExtractRegistryStringValue(key.GetValue(valueName));
            if (string.IsNullOrWhiteSpace(value))
            {
                continue;
            }

            foreach (var identifier in identifiers)
            {
                if (value.Equals(identifier, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
        }

        foreach (var valueName in valueNames)
        {
            if (candidateValueNames.Contains(valueName))
            {
                continue;
            }

            var value = ExtractRegistryStringValue(key.GetValue(valueName));
            if (string.IsNullOrWhiteSpace(value))
            {
                continue;
            }

            foreach (var identifier in identifiers)
            {
                if (value.Equals(identifier, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
        }

        return false;
    }

    private static string? ExtractRegistryStringValue(object? registryValue)
    {
        switch (registryValue)
        {
            case string text:
                return text.TrimEnd('\0');
            case byte[] bytes when bytes.Length > 0:
                return Encoding.Unicode.GetString(bytes).TrimEnd('\0');
            default:
                return null;
        }
    }

    private string? ResolveDefaultProfileName(string version)
    {
        try
        {
            using var outlookKey = Registry.CurrentUser.OpenSubKey(
                $@"Software\Microsoft\Office\{version}\Outlook",
                writable: false);

            var defaultProfile = outlookKey?.GetValue("DefaultProfile") as string;
            if (!string.IsNullOrWhiteSpace(defaultProfile))
            {
                return defaultProfile;
            }

            using var profilesKey = Registry.CurrentUser.OpenSubKey(
                $@"Software\Microsoft\Office\{version}\Outlook\Profiles",
                writable: false);

            var profileNames = profilesKey?.GetSubKeyNames();
            if (profileNames != null && profileNames.Length > 0)
            {
                return profileNames[0];
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Errore risoluzione profilo Outlook: {ex.Message}");
        }

        return null;
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
