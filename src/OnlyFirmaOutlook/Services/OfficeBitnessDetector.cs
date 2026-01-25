using System.IO;
using Microsoft.Win32;

namespace OnlyFirmaOutlook.Services;





public static class OfficeBitnessDetector
{
    private static readonly LoggingService _logger = LoggingService.Instance;

    public enum OfficeBitness
    {
        Unknown,
        x86,
        x64
    }

    
    
    
    public static OfficeBitness DetectOfficeBitness()
    {
        _logger.Log("Rilevamento bitness Office in corso...");

        
        var outlookBitness = DetectFromOutlookKey();
        if (outlookBitness != OfficeBitness.Unknown)
        {
            _logger.Log($"Bitness rilevata da chiave Outlook: {outlookBitness}");
            return outlookBitness;
        }

        
        var clickToRunBitness = DetectFromClickToRun();
        if (clickToRunBitness != OfficeBitness.Unknown)
        {
            _logger.Log($"Bitness rilevata da ClickToRun: {clickToRunBitness}");
            return clickToRunBitness;
        }

        
        var msiBitness = DetectFromMsiInstallation();
        if (msiBitness != OfficeBitness.Unknown)
        {
            _logger.Log($"Bitness rilevata da MSI: {msiBitness}");
            return msiBitness;
        }

        
        var wordExeBitness = DetectFromWordExecutable();
        if (wordExeBitness != OfficeBitness.Unknown)
        {
            _logger.Log($"Bitness rilevata da eseguibile Word: {wordExeBitness}");
            return wordExeBitness;
        }

        _logger.LogWarning("Impossibile determinare bitness Office. Utilizzo x64 come default.");
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
                _logger.LogWarning($"Errore lettura chiave Outlook {version}: {ex.Message}");
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
            _logger.LogWarning($"Errore lettura chiave ClickToRun: {ex.Message}");
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
                _logger.LogWarning($"Errore lettura chiave MSI {version}: {ex.Message}");
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
                _logger.Log($"Trovato Word in: {path}");
                return path.Contains("Program Files (x86)", StringComparison.OrdinalIgnoreCase)
                    ? OfficeBitness.x86
                    : OfficeBitness.x64;
            }
        }

        return OfficeBitness.Unknown;
    }

    
    
    
    public static bool IsWordInstalled()
    {
        _logger.Log("Verifica installazione Word...");

        try
        {
            
            var wordType = Type.GetTypeFromProgID("Word.Application");
            var installed = wordType != null;
            _logger.Log($"Word installato: {installed}");
            return installed;
        }
        catch (Exception ex)
        {
            _logger.LogError("Errore verifica Word", ex);
            return false;
        }
    }

    
    
    
    public static bool IsOutlookInstalled()
    {
        _logger.Log("Verifica installazione Outlook...");

        try
        {
            var outlookType = Type.GetTypeFromProgID("Outlook.Application");
            var installed = outlookType != null;
            _logger.Log($"Outlook installato: {installed}");
            return installed;
        }
        catch (Exception ex)
        {
            _logger.LogError("Errore verifica Outlook", ex);
            return false;
        }
    }
}
