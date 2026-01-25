using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace OnlyFirmaOutlook.Launcher;

/// <summary>
/// Bootstrapper/Launcher che rileva la bitness di Office installato
/// e avvia l'eseguibile corretto (x86 o x64).
/// </summary>
internal class Program
{
    private const string AppName = "OnlyFirmaOutlook";
    private const string ExeName = "OnlyFirmaOutlook.exe";

    [STAThread]
    static int Main(string[] args)
    {
        try
        {
            var baseDir = AppContext.BaseDirectory;
            var bitness = DetectOfficeBitness();

            string targetDir;
            string fallbackDir;

            if (bitness == OfficeBitness.x86)
            {
                targetDir = Path.Combine(baseDir, "win-x86");
                fallbackDir = Path.Combine(baseDir, "win-x64");
                Console.WriteLine($"[Launcher] Rilevato Office 32-bit, avvio versione x86...");
            }
            else
            {
                // Default a x64 se non rilevabile o se Office è 64-bit
                targetDir = Path.Combine(baseDir, "win-x64");
                fallbackDir = Path.Combine(baseDir, "win-x86");

                if (bitness == OfficeBitness.x64)
                {
                    Console.WriteLine($"[Launcher] Rilevato Office 64-bit, avvio versione x64...");
                }
                else
                {
                    Console.WriteLine($"[Launcher] Bitness Office non determinata, utilizzo x64 come default...");
                }
            }

            var exePath = Path.Combine(targetDir, ExeName);

            // Se l'eseguibile target non esiste, prova il fallback
            if (!File.Exists(exePath))
            {
                Console.WriteLine($"[Launcher] Eseguibile non trovato in {targetDir}");
                exePath = Path.Combine(fallbackDir, ExeName);

                if (!File.Exists(exePath))
                {
                    ShowError(
                        $"Impossibile trovare {ExeName}.\n\n" +
                        $"Cercato in:\n- {targetDir}\n- {fallbackDir}\n\n" +
                        "Assicurarsi che l'applicazione sia stata pubblicata correttamente.");
                    return 1;
                }

                Console.WriteLine($"[Launcher] Utilizzo fallback: {exePath}");
            }

            Console.WriteLine($"[Launcher] Avvio: {exePath}");

            // Avvia l'applicazione
            var startInfo = new ProcessStartInfo
            {
                FileName = exePath,
                WorkingDirectory = Path.GetDirectoryName(exePath),
                UseShellExecute = false
            };

            // Passa eventuali argomenti
            foreach (var arg in args)
            {
                startInfo.ArgumentList.Add(arg);
            }

            var process = Process.Start(startInfo);

            if (process == null)
            {
                ShowError($"Impossibile avviare {ExeName}.");
                return 1;
            }

            return 0;
        }
        catch (Exception ex)
        {
            ShowError($"Errore durante l'avvio:\n\n{ex.Message}");
            return 1;
        }
    }

    private enum OfficeBitness
    {
        Unknown,
        x86,
        x64
    }

    private static OfficeBitness DetectOfficeBitness()
    {
        Console.WriteLine("[Launcher] Rilevamento bitness Office...");

        // Metodo 1: Chiave Outlook
        var outlookBitness = DetectFromOutlookKey();
        if (outlookBitness != OfficeBitness.Unknown)
        {
            return outlookBitness;
        }

        // Metodo 2: Chiave ClickToRun
        var clickToRunBitness = DetectFromClickToRun();
        if (clickToRunBitness != OfficeBitness.Unknown)
        {
            return clickToRunBitness;
        }

        // Metodo 3: Chiave Word MSI
        var msiBitness = DetectFromMsiInstallation();
        if (msiBitness != OfficeBitness.Unknown)
        {
            return msiBitness;
        }

        // Metodo 4: Verifica esistenza file Word
        var wordExeBitness = DetectFromWordExecutable();
        if (wordExeBitness != OfficeBitness.Unknown)
        {
            return wordExeBitness;
        }

        return OfficeBitness.Unknown;
    }

    private static OfficeBitness DetectFromOutlookKey()
    {
        string[] versions = { "16.0", "15.0", "14.0" };

        foreach (var version in versions)
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
                        Console.WriteLine($"[Launcher] Trovata chiave Outlook {version}, Bitness={bitness}");
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
                        Console.WriteLine($"[Launcher] Trovata chiave Outlook {version} (WOW64), Bitness={bitness}");
                        return bitness.Equals("x64", StringComparison.OrdinalIgnoreCase)
                            ? OfficeBitness.x64
                            : OfficeBitness.x86;
                    }
                    return OfficeBitness.x86;
                }
            }
            catch
            {
                // Ignora errori di accesso al registro
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
                    Console.WriteLine($"[Launcher] Trovata chiave ClickToRun, Platform={platform}");
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
                    Console.WriteLine($"[Launcher] Trovata chiave ClickToRun (WOW64), Platform={platform}");
                    return platform.Equals("x64", StringComparison.OrdinalIgnoreCase)
                        ? OfficeBitness.x64
                        : OfficeBitness.x86;
                }
                return OfficeBitness.x86;
            }
        }
        catch
        {
            // Ignora errori di accesso al registro
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
                        Console.WriteLine($"[Launcher] Trovato Word MSI {version}: {path}");
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
                    Console.WriteLine($"[Launcher] Trovato Word MSI {version} in WOW64 (32-bit)");
                    return OfficeBitness.x86;
                }
            }
            catch
            {
                // Ignora errori di accesso al registro
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
                Console.WriteLine($"[Launcher] Trovato Word: {path}");
                return path.Contains("Program Files (x86)", StringComparison.OrdinalIgnoreCase)
                    ? OfficeBitness.x86
                    : OfficeBitness.x64;
            }
        }

        return OfficeBitness.Unknown;
    }

    private static void ShowError(string message)
    {
        Console.Error.WriteLine($"[Launcher] ERRORE: {message}");

        // Mostra anche un MessageBox per l'utente
        MessageBox(IntPtr.Zero, message, $"{AppName} - Errore", 0x10); // MB_ICONERROR
    }

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int MessageBox(IntPtr hWnd, string text, string caption, uint type);
}
