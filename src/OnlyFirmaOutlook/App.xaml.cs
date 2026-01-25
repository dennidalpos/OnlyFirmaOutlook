using System.Windows;
using OnlyFirmaOutlook.Services;
using Application = System.Windows.Application;

namespace OnlyFirmaOutlook;




public partial class App : Application
{
    private readonly TempFileManager _tempFileManager;
    private readonly LoggingService _loggingService;

    public App()
    {
        _loggingService = LoggingService.Instance;
        _tempFileManager = TempFileManager.Instance;
    }

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        _loggingService.Log("Applicazione avviata");
        _loggingService.Log($"Directory base: {AppContext.BaseDirectory}");
        _loggingService.Log($"Cartella temporanea sessione: {_tempFileManager.SessionTempFolder}");
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _loggingService.Log("Chiusura applicazione in corso...");

        
        _tempFileManager.CleanupSessionFolder();

        _loggingService.Log("Applicazione chiusa");
        _loggingService.Dispose();

        base.OnExit(e);
    }
}
