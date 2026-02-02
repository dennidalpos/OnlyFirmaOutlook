using System.Windows;
using OnlyFirmaOutlook.Services;
using Application = System.Windows.Application;

namespace OnlyFirmaOutlook;




public partial class App : Application
{
    private readonly TempFileManager _tempFileManager;
    private readonly LoggingService _loggingService;
    private readonly WordEditorService _wordEditorService;

    public App()
    {
        _loggingService = LoggingService.Instance;
        _tempFileManager = TempFileManager.Instance;
        _wordEditorService = new WordEditorService();

        OfficeBitnessDetector.LogInfo = _loggingService.Log;
        OfficeBitnessDetector.LogWarning = _loggingService.LogWarning;
        OfficeBitnessDetector.LogError = _loggingService.LogError;
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
        _wordEditorService.CleanupAllEditorFolders();

        _loggingService.Log("Applicazione chiusa");
        _loggingService.Dispose();

        base.OnExit(e);
    }
}
