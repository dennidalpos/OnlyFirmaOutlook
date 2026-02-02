using System.Diagnostics;
using System.Runtime.InteropServices;
using OnlyFirmaOutlook.Services;

namespace OnlyFirmaOutlook.Launcher;





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
            ConfigureLogging();
            var bitness = OfficeBitnessDetector.DetectOfficeBitness();

            string targetDir;
            string fallbackDir;

            if (bitness == OfficeBitnessDetector.OfficeBitness.x86)
            {
                targetDir = Path.Combine(baseDir, "win-x86");
                fallbackDir = Path.Combine(baseDir, "win-x64");
                Console.WriteLine($"[Launcher] Rilevato Office 32-bit, avvio versione x86...");
            }
            else
            {
                
                targetDir = Path.Combine(baseDir, "win-x64");
                fallbackDir = Path.Combine(baseDir, "win-x86");

                if (bitness == OfficeBitnessDetector.OfficeBitness.x64)
                {
                    Console.WriteLine($"[Launcher] Rilevato Office 64-bit, avvio versione x64...");
                }
                else
                {
                    Console.WriteLine($"[Launcher] Bitness Office non determinata, utilizzo x64 come default...");
                }
            }

            var exePath = Path.Combine(targetDir, ExeName);

            
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

            
            var startInfo = new ProcessStartInfo
            {
                FileName = exePath,
                WorkingDirectory = Path.GetDirectoryName(exePath),
                UseShellExecute = false
            };

            
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

    private static void ConfigureLogging()
    {
        OfficeBitnessDetector.LogInfo = message => Console.WriteLine($"[Launcher] {message}");
        OfficeBitnessDetector.LogWarning = message => Console.WriteLine($"[Launcher] AVVISO: {message}");
        OfficeBitnessDetector.LogError = (message, ex) =>
            Console.WriteLine($"[Launcher] ERRORE: {message} - {ex.Message}");
    }

    private static void ShowError(string message)
    {
        Console.Error.WriteLine($"[Launcher] ERRORE: {message}");

        
        MessageBox(IntPtr.Zero, message, $"{AppName} - Errore", 0x10); 
    }

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int MessageBox(IntPtr hWnd, string text, string caption, uint type);
}
