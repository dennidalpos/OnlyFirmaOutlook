using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Threading;
using OnlyFirmaOutlook.Models;
using OnlyFirmaOutlook.Services;
using MessageBox = System.Windows.MessageBox;

namespace OnlyFirmaOutlook.Views;





public partial class MainWindow : Window
{
    private readonly LoggingService _logger;
    private readonly TempFileManager _tempFileManager;
    private readonly PresetService _presetService;
    private readonly OutlookAccountService _outlookAccountService;
    private readonly SignatureRepository _signatureRepository;
    private readonly WordConversionService _wordConversionService;
    private readonly WordEditorService _wordEditorService;
    private readonly SignatureWorkflowService _signatureWorkflowService;

    private List<PresetFile> _presets = new();
    private List<OutlookAccount> _accounts = new();
    private List<SignatureInfo> _existingSignatures = new();
    private List<BackupInfo> _existingBackups = new();

    private string? _selectedFilePath;
    private EditorState? _currentEditorState;
    private bool _isOutlookAvailable;
    private bool _isFolderWritable;

    
    private FileSystemWatcher? _fileWatcher;
    private DispatcherTimer? _wordCheckTimer;
    private DateTime _lastFileModifiedTime;
    private bool _isWordOpen;
    private GuideWindow? _guideWindow;

    public MainWindow()
    {
        InitializeComponent();

        
        _logger = LoggingService.Instance;
        _tempFileManager = TempFileManager.Instance;
        _presetService = new PresetService();
        _outlookAccountService = new OutlookAccountService();
        _signatureRepository = new SignatureRepository();
        _wordConversionService = new WordConversionService();
        _wordEditorService = new WordEditorService();
        _signatureWorkflowService = new SignatureWorkflowService(_signatureRepository, _wordConversionService);

        
        _logger.LogAdded += OnLogAdded;

        UpdateHeaderInfo();

        
        LogTextBox.Text = _logger.GetFullLog();
        ScrollLogToEnd();

        
        Loaded += MainWindow_Loaded;
    }

    private void UpdateHeaderInfo()
    {
        var version = Assembly.GetExecutingAssembly().GetName().Version;
        var displayVersion = version == null
            ? "N/D"
            : $"{version.Major}.{version.Minor}.{version.Build}";

        AppVersionText.Text = $"Versione {displayVersion}";

        var hostName = Environment.MachineName;
        var userName = Environment.UserName;
        var ipAddress = GetLocalIpAddress();

        UserHostInfoText.Text = $"Hostname: {hostName} | IP: {ipAddress} | Utente: {userName}";
    }

    private static string GetLocalIpAddress()
    {
        try
        {
            var addresses = Dns.GetHostAddresses(Dns.GetHostName());
            foreach (var address in addresses)
            {
                if (address.AddressFamily == AddressFamily.InterNetwork && !IPAddress.IsLoopback(address))
                {
                    return address.ToString();
                }
            }
        }
        catch
        {
            return "N/D";
        }

        return "N/D";
    }

    private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        _logger.Log("Inizializzazione interfaccia...");

        
        _tempFileManager.CleanupOrphanedFolders();
        _wordEditorService.CleanupOrphanedEditorFolders();

        
        LoadPresets();

        
        await InitializeAsync();

        _logger.Log("Interfaccia pronta");
    }

    private async Task InitializeAsync()
    {
        SetBusy(true, "Rilevamento configurazione Office...");

        try
        {
            
            if (!OfficeBitnessDetector.IsWordInstalled())
            {
                MessageBox.Show(
                    "Microsoft Word non risulta installato.\n\n" +
                    "Word è necessario per la conversione dei documenti.",
                    "Word non trovato",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }

            
            var accountResult = await Task.Run(() => _outlookAccountService.LoadAccounts());

            _isOutlookAvailable = accountResult.OutlookAvailable;
            _accounts = accountResult.Accounts;

            
            ConfigureOutlookUI();

            
            SetDefaultDestinationFolder();

            
            RefreshExistingSignatures();

            
            RefreshBackups();
        }
        finally
        {
            SetBusy(false);
        }
    }

    private void LoadPresets()
    {
        _presets = _presetService.LoadPresets();

        if (_presets.Count > 0)
        {
            PresetListBox.ItemsSource = _presets;
            PresetListBox.DisplayMemberPath = "DisplayName";
            PresetListBox.SelectionChanged += PresetListBox_SelectionChanged;
            NoPresetsText.Visibility = Visibility.Collapsed;
        }
        else
        {
            PresetListBox.ItemsSource = null;
            NoPresetsText.Visibility = Visibility.Visible;
        }
    }

    private void ConfigureOutlookUI()
    {
        if (_isOutlookAvailable && _accounts.Count > 0)
        {
            
            OutlookWarningBorder.Visibility = Visibility.Collapsed;
            AccountLabel.Visibility = Visibility.Visible;
            AccountComboBox.Visibility = Visibility.Visible;
            var view = CollectionViewSource.GetDefaultView(_accounts);
            view.GroupDescriptions.Clear();
            view.GroupDescriptions.Add(new PropertyGroupDescription(nameof(OutlookAccount.GroupLabel)));
            AccountComboBox.ItemsSource = view;
            AccountComboBox.DisplayMemberPath = "DisplayText";

            IdentifierLabel.Visibility = Visibility.Collapsed;
            IdentifierTextBox.Visibility = Visibility.Collapsed;
            IdentifierHint.Visibility = Visibility.Collapsed;

            
            if (_accounts.Count > 0)
            {
                AccountComboBox.SelectedIndex = 0;
            }
        }
        else if (_isOutlookAvailable && _accounts.Count == 0)
        {
            
            OutlookWarningBorder.Visibility = Visibility.Visible;
            OutlookWarningText.Text = "Outlook è installato ma non sono configurati account. " +
                "Puoi comunque creare la firma e copiarla manualmente.";

            AccountLabel.Visibility = Visibility.Collapsed;
            AccountComboBox.Visibility = Visibility.Collapsed;

            IdentifierLabel.Visibility = Visibility.Visible;
            IdentifierTextBox.Visibility = Visibility.Visible;
            IdentifierHint.Visibility = Visibility.Visible;
        }
        else
        {
            
            OutlookWarningBorder.Visibility = Visibility.Visible;
            OutlookWarningText.Text = "Outlook non è installato. La firma verrà salvata in una cartella locale. " +
                "Potrai poi copiarla manualmente in %APPDATA%\\Microsoft\\Signatures.";

            AccountLabel.Visibility = Visibility.Collapsed;
            AccountComboBox.Visibility = Visibility.Collapsed;

            IdentifierLabel.Visibility = Visibility.Visible;
            IdentifierTextBox.Visibility = Visibility.Visible;
            IdentifierHint.Visibility = Visibility.Visible;
        }

    }

    private void SetDefaultDestinationFolder()
    {
        string defaultFolder;

        if (_isOutlookAvailable)
        {
            defaultFolder = SignatureRepository.GetDefaultOutlookSignaturesFolder();
        }
        else
        {
            defaultFolder = SignatureRepository.GetAlternativeOutputFolder();
        }

        DestinationFolderTextBox.Text = defaultFolder;
        ValidateDestinationFolder(defaultFolder);
    }

    private void ValidateDestinationFolder(string folderPath)
    {
        _isFolderWritable = _signatureRepository.CanWriteToFolder(folderPath);

        if (_isFolderWritable)
        {
            FolderWritableText.Text = "Cartella scrivibile";
            FolderWritableText.Foreground = System.Windows.Media.Brushes.Green;
        }
        else
        {
            FolderWritableText.Text = "Cartella non scrivibile - selezionare un'altra cartella";
            FolderWritableText.Foreground = System.Windows.Media.Brushes.Red;
        }

        UpdateConvertButtonState();
    }

    private void UpdateConvertButtonState()
    {
        var hasFile = !string.IsNullOrEmpty(_selectedFilePath) && File.Exists(_selectedFilePath);
        var hasSignatureName = !string.IsNullOrWhiteSpace(SignatureNameTextBox.Text);
        var hasDestination = !string.IsNullOrWhiteSpace(DestinationFolderTextBox.Text);
        var isDocumentReady = _currentEditorState?.IsReadyForConversion ?? false;

        
        ConvertButton.IsEnabled = hasFile && hasSignatureName && hasDestination && _isFolderWritable && isDocumentReady;

        
        UpdateEditStatusDisplay();

        UpdateFinalSignatureName();

        
        UpdateStepHighlighting();

        
        CheckOverwriteWarning();
    }

    
    
    
    private void UpdateStepHighlighting()
    {
        var hasSignatureSelected = _currentEditorState != null;
        var hasSignatureName = !string.IsNullOrWhiteSpace(SignatureNameTextBox.Text);
        var hasDestination = _isFolderWritable;
        var isDocumentReady = _currentEditorState?.IsReadyForConversion ?? false;

        if (Resources.Contains("StepGroupBoxStyle"))
        {
            Step3Group.Style = (Style)Resources["StepGroupBoxStyle"];
            Step5Group.Style = (Style)Resources["StepGroupBoxStyle"];
            Step6Group.Style = (Style)Resources["StepGroupBoxStyle"];
            Step7Group.Style = (Style)Resources["StepGroupBoxStyle"];
        }

        
        if (!hasSignatureSelected)
        {
            SetStepStyle(Step1Group, StepState.Current);
            SetStepStyle(Step2Group, StepState.Pending);
            SetStepStyle(Step4Group, StepState.Pending);
            return;
        }

        SetStepStyle(Step1Group, StepState.Completed);

        
        if (!hasSignatureName)
        {
            SetStepStyle(Step2Group, StepState.Current);
            SetStepStyle(Step4Group, StepState.Pending);
            return;
        }

        SetStepStyle(Step2Group, StepState.Completed);

        
        if (!isDocumentReady)
        {
            SetStepStyle(Step4Group, StepState.Current);
            return;
        }

        SetStepStyle(Step4Group, StepState.Completed);

        
    }

    private enum StepState { Pending, Current, Completed }

    private void SetStepStyle(System.Windows.Controls.GroupBox groupBox, StepState state)
    {
        var styleName = state switch
        {
            StepState.Completed => "CompletedStepStyle",
            StepState.Current => "CurrentStepStyle",
            _ => "PendingStepStyle"
        };

        if (Resources.Contains(styleName))
        {
            groupBox.Style = (Style)Resources[styleName];
        }
    }

    
    
    
    private void UpdateEditStatusDisplay()
    {
        if (_currentEditorState == null)
        {
            EditStatusText.Text = "Nessuna firma selezionata";
            EditStatusText.Foreground = System.Windows.Media.Brushes.Gray;
            EditSignatureButton.IsEnabled = false;
            return;
        }

        EditStatusText.Text = _currentEditorState.GetStatusText();

        if (_currentEditorState.IsReadyForConversion)
        {
            EditStatusText.Foreground = System.Windows.Media.Brushes.Green;
        }
        else
        {
            EditStatusText.Foreground = System.Windows.Media.Brushes.OrangeRed;
        }

        EditSignatureButton.IsEnabled = true;
    }

    private async void ConvertButton_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_selectedFilePath))
        {
            MessageBox.Show("Seleziona un documento Word.", "Attenzione", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        
        if (_currentEditorState == null || !_currentEditorState.IsReadyForConversion)
        {
            MessageBox.Show(
                "Il documento deve essere modificato e salvato nell'editor prima di poter essere convertito.\n\n" +
                "Clicca su 'Modifica firma' per aprire l'editor.",
                "Documento non pronto",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        var baseName = SignatureNameTextBox.Text.Trim();
        if (string.IsNullOrEmpty(baseName))
        {
            MessageBox.Show("Inserisci un nome per la firma.", "Attenzione", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        
        var finalSignatureName = _signatureWorkflowService.BuildFinalSignatureName(baseName, GetCurrentSignatureIdentifier());
        var destinationFolder = DestinationFolderTextBox.Text;

        if (!_isFolderWritable)
        {
            MessageBox.Show(
                "La cartella di destinazione non è scrivibile. Seleziona un'altra cartella.",
                "Cartella non scrivibile",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        _signatureWorkflowService.CreateBackupIfNeeded(destinationFolder);
        RefreshBackups();

        
        if (_signatureWorkflowService.SignatureExists(destinationFolder, finalSignatureName))
        {
            var result = MessageBox.Show(
                $"Esiste già una firma con il nome '{finalSignatureName}'.\n\nVuoi sovrascriverla?",
                "Firma esistente",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result != MessageBoxResult.Yes)
            {
                return;
            }

            
            _signatureWorkflowService.DeleteExistingSignatureFiles(destinationFolder, finalSignatureName);
        }

        SetBusy(true, "Conversione in corso...");

        try
        {
            var useFilteredHtml = FilteredHtmlRadio.IsChecked == true;

            var conversionResult = await Task.Run(() =>
            {
                return _signatureWorkflowService.ConvertDocument(
                    _selectedFilePath,
                    destinationFolder,
                    finalSignatureName,
                    useFilteredHtml);
            });

            if (conversionResult.Success)
            {
                _logger.Log("Conversione completata con successo!");

                
                if (_currentEditorState != null)
                {
                    _wordEditorService.CleanupEditorTempFolder(_currentEditorState.EditorSessionId);
                    _currentEditorState = null;
                }

                
                _selectedFilePath = null;
                UpdateConvertButtonState();


                RefreshExistingSignatures();

                MessageBox.Show(
                    $"Firma '{finalSignatureName}' creata con successo!\n\n" +
                    $"File creati:\n" +
                    $"- {Path.GetFileName(conversionResult.HtmFilePath)}\n" +
                    $"- {Path.GetFileName(conversionResult.RtfFilePath)}\n" +
                    $"- {Path.GetFileName(conversionResult.TxtFilePath)}" +
                    (conversionResult.AssetsFolderPath != null
                        ? $"\n- {Path.GetFileName(conversionResult.AssetsFolderPath)}/"
                        : string.Empty),
                    "Conversione completata",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                ResetUiForNewSignature();
            }
            else
            {
                MessageBox.Show(
                    $"Errore durante la conversione:\n\n{conversionResult.ErrorMessage}",
                    "Errore conversione",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }
        catch (Exception ex)
        {
            _logger.LogError("Errore durante la conversione", ex);
            MessageBox.Show(
                $"Errore durante la conversione:\n\n{ex.Message}",
                "Errore",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            SetBusy(false);
        }
    }


    protected override void OnClosed(EventArgs e)
    {
        if (_guideWindow != null)
        {
            _guideWindow.Closed -= GuideWindow_Closed;
            _guideWindow.Close();
            _guideWindow = null;
        }

        StopFileWatcher();
        StopWordCheckTimer();

        _logger.LogAdded -= OnLogAdded;
        base.OnClosed(e);
    }
}
