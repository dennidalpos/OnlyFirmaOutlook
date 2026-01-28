using System.IO;
using Microsoft.Win32;

namespace OnlyFirmaOutlook.Services;

public static class OfficeBitnessDetector
{
    public static Action<string>? LogInfo { get; set; }
    public static Action<string>? LogWarning { get; set; }
    public static Action<string, Exception>? LogError { get; set; }

    public enum OfficeBitness
    {
        Unknown,
        x86,
        x64
    }

    public static OfficeBitness DetectOfficeBitness()
    {
        Log("Rilevamento bitness Office in corso...");

        var outlookBitness = DetectFromOutlookKey();
        if (outlookBitness != OfficeBitness.Unknown)
        {
            Log($"Bitness rilevata da chiave Outlook: {outlookBitness}");
            return outlookBitness;
        }

        var clickToRunBitness = DetectFromClickToRun();
        if (clickToRunBitness != OfficeBitness.Unknown)
        {
            Log($"Bitness rilevata da ClickToRun: {clickToRunBitness}");
            return clickToRunBitness;
        }

        var msiBitness = DetectFromMsiInstallation();
        if (msiBitness != OfficeBitness.Unknown)
        {
            Log($"Bitness rilevata da MSI: {msiBitness}");
            return msiBitness;
        }

        var wordExeBitness = DetectFromWordExecutable();
        if (wordExeBitness != OfficeBitness.Unknown)
        {
            Log($"Bitness rilevata da eseguibile Word: {wordExeBitness}");
            return wordExeBitness;
        }

        Warn("Impossibile determinare bitness Office. Utilizzo x64 come default.");
        return OfficeBitness.Unknown;
    }

    private static OfficeBitness DetectFromOutlookKey()
    {
        string[] outlookVersions = { "16.0", "15.0", "14.0" };

        foreach (var version in outlookVersions)
        {
            try
            {
                using var key = Registry.LocalMachine.OpenSubKey(
                    $@"SOFTWARE\Microsoft\Office\{version}\Outlook", false);
                if (key != null)
                {
                    var bitness = key.GetValue("Bitness") as string;
                    if (!string.IsNullOrEmpty(bitness))
                    {
                        return bitness.Equals("x64", StringComparison.OrdinalIgnoreCase)
                            ? OfficeBitness.x64
                            : OfficeBitness.x86;
                    }
                }

                using var key32 = Registry.LocalMachine.OpenSubKey(
                    $@"SOFTWARE\WOW6432Node\Microsoft\Office\{version}\Outlook", false);
                if (key32 != null)
                {
                    var bitness = key32.GetValue("Bitness") as string;
                    if (!string.IsNullOrEmpty(bitness))
                    {
                        return bitness.Equals("x64", StringComparison.OrdinalIgnoreCase)
                            ? OfficeBitness.x64
                            : OfficeBitness.x86;
                    }

                    return OfficeBitness.x86;
                }
            }
            catch (Exception ex)
            {
                Warn($"Errore lettura chiave Outlook {version}: {ex.Message}");
            }
        }

        return OfficeBitness.Unknown;
    }

    private static OfficeBitness DetectFromClickToRun()
    {
        try
        {
            using var key = Registry.LocalMachine.OpenSubKey(
                @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration", false);
            if (key != null)
            {
                var platform = key.GetValue("Platform") as string;
                if (!string.IsNullOrEmpty(platform))
                {
                    return platform.Equals("x64", StringComparison.OrdinalIgnoreCase)
                        ? OfficeBitness.x64
                        : OfficeBitness.x86;
                }
            }

            using var key32 = Registry.LocalMachine.OpenSubKey(
                @"SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration", false);
            if (key32 != null)
            {
                var platform = key32.GetValue("Platform") as string;
                if (!string.IsNullOrEmpty(platform))
                {
                    return platform.Equals("x64", StringComparison.OrdinalIgnoreCase)
                        ? OfficeBitness.x64
                        : OfficeBitness.x86;
                }
                return OfficeBitness.x86;
            }
        }
        catch (Exception ex)
        {
            Warn($"Errore lettura chiave ClickToRun: {ex.Message}");
        }

        return OfficeBitness.Unknown;
    }

    private static OfficeBitness DetectFromMsiInstallation()
    {
        string[] versions = { "16.0", "15.0", "14.0" };

        foreach (var version in versions)
        {
            try
            {
                using var key = Registry.LocalMachine.OpenSubKey(
                    $@"SOFTWARE\Microsoft\Office\{version}\Word\InstallRoot", false);
                if (key != null)
                {
                    var path = key.GetValue("Path") as string;
                    if (!string.IsNullOrEmpty(path))
                    {
                        if (path.Contains("Program Files (x86)", StringComparison.OrdinalIgnoreCase))
                        {
                            return OfficeBitness.x86;
                        }
                        if (path.Contains("Program Files", StringComparison.OrdinalIgnoreCase))
                        {
                            return OfficeBitness.x64;
                        }
                    }
                }

                using var key32 = Registry.LocalMachine.OpenSubKey(
                    $@"SOFTWARE\WOW6432Node\Microsoft\Office\{version}\Word\InstallRoot", false);
                if (key32 != null)
                {
                    return OfficeBitness.x86;
                }
            }
            catch (Exception ex)
            {
                Warn($"Errore lettura chiave MSI {version}: {ex.Message}");
            }
        }

        return OfficeBitness.Unknown;
    }

    private static OfficeBitness DetectFromWordExecutable()
    {
        string[] possiblePaths =
        {
            @"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
            @"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE",
            @"C:\Program Files\Microsoft Office\Office16\WINWORD.EXE",
            @"C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE",
            @"C:\Program Files\Microsoft Office\Office15\WINWORD.EXE",
            @"C:\Program Files (x86)\Microsoft Office\Office15\WINWORD.EXE",
        };

        foreach (var path in possiblePaths)
        {
            if (File.Exists(path))
            {
                Log($"Trovato Word in: {path}");
                return path.Contains("Program Files (x86)", StringComparison.OrdinalIgnoreCase)
                    ? OfficeBitness.x86
                    : OfficeBitness.x64;
            }
        }

        return OfficeBitness.Unknown;
    }

    public static bool IsWordInstalled()
    {
        Log("Verifica installazione Word...");

        try
        {
            var wordType = Type.GetTypeFromProgID("Word.Application");
            var installed = wordType != null;
            Log($"Word installato: {installed}");
            return installed;
        }
        catch (Exception ex)
        {
            Error("Errore verifica Word", ex);
            return false;
        }
    }

    public static bool IsOutlookInstalled()
    {
        Log("Verifica installazione Outlook...");

        try
        {
            var outlookType = Type.GetTypeFromProgID("Outlook.Application");
            var installed = outlookType != null;
            Log($"Outlook installato: {installed}");
            return installed;
        }
        catch (Exception ex)
        {
            Error("Errore verifica Outlook", ex);
            return false;
        }
    }

    private static void Log(string message)
    {
        LogInfo?.Invoke(message);
    }

    private static void Warn(string message)
    {
        LogWarning?.Invoke(message);
    }

    private static void Error(string message, Exception ex)
    {
        LogError?.Invoke(message, ex);
    }
}
