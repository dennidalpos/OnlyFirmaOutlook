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
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using MessageBox = System.Windows.MessageBox;
using Clipboard = System.Windows.Clipboard;

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

    private void RefreshExistingSignatures()
    {
        var folderPath = DestinationFolderTextBox.Text;

        if (string.IsNullOrEmpty(folderPath))
        {
            _existingSignatures = new List<SignatureInfo>();
        }
        else
        {
            _existingSignatures = _signatureRepository.GetSignatures(folderPath);
        }

        if (_existingSignatures.Count > 0)
        {
            ExistingSignaturesListBox.ItemsSource = _existingSignatures;
            ExistingSignaturesListBox.DisplayMemberPath = "DisplayInfo";
            NoSignaturesText.Visibility = Visibility.Collapsed;
        }
        else
        {
            ExistingSignaturesListBox.ItemsSource = null;
            NoSignaturesText.Visibility = Visibility.Visible;
        }

        DeleteSignatureButton.IsEnabled = false;
    }

    private void RefreshBackups()
    {
        var backupsFolder = SignatureRepository.GetDefaultOutlookSignaturesFolder();
        BackupFolderText.Text = $"Cartella backup: {backupsFolder}";

        _existingBackups = _signatureRepository.GetBackups(backupsFolder);

        if (_existingBackups.Count > 0)
        {
            BackupsListBox.ItemsSource = _existingBackups;
            BackupsListBox.DisplayMemberPath = "DisplayInfo";
            NoBackupsText.Visibility = Visibility.Collapsed;
        }
        else
        {
            BackupsListBox.ItemsSource = null;
            NoBackupsText.Visibility = Visibility.Visible;
        }

        UpdateBackupButtons();
    }

    private void UpdateBackupButtons()
    {
        var hasSelection = BackupsListBox.SelectedItem != null;
        RestoreBackupButton.IsEnabled = hasSelection;
        DeleteBackupButton.IsEnabled = hasSelection;
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

    
    
    
    private void CheckOverwriteWarning()
    {
        var baseName = SignatureNameTextBox.Text?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(baseName))
        {
            OverwriteWarningBorder.Visibility = Visibility.Collapsed;
            return;
        }

        string? identifier = null;
        if (_isOutlookAvailable && AccountComboBox.SelectedItem is OutlookAccount account)
        {
            identifier = account.DisplayText;
        }
        else if (!string.IsNullOrWhiteSpace(IdentifierTextBox.Text))
        {
            identifier = IdentifierTextBox.Text.Trim();
        }

        var finalName = _signatureWorkflowService.BuildFinalSignatureName(baseName, identifier);
        var destinationFolder = DestinationFolderTextBox.Text;



        if (!string.IsNullOrEmpty(destinationFolder) &&
            _signatureWorkflowService.SignatureExists(destinationFolder, finalName))
        {
            OverwriteWarningText.Text = $"La firma '{finalName}' esiste già e verrà sovrascritta!";
            OverwriteWarningBorder.Visibility = Visibility.Visible;
        }
        else
        {
            OverwriteWarningBorder.Visibility = Visibility.Collapsed;
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

    private void UpdateFinalSignatureName()
    {
        var baseName = SignatureNameTextBox.Text?.Trim() ?? string.Empty;

        if (string.IsNullOrEmpty(baseName))
        {
            FinalNameBorder.Visibility = Visibility.Collapsed;
            return;
        }

        string? identifier = null;

        if (_isOutlookAvailable && AccountComboBox.SelectedItem is OutlookAccount account)
        {
            identifier = account.DisplayText;
        }
        else if (!string.IsNullOrWhiteSpace(IdentifierTextBox.Text))
        {
            identifier = IdentifierTextBox.Text.Trim();
        }

        var finalName = _signatureWorkflowService.BuildFinalSignatureName(baseName, identifier);

        FinalSignatureNameText.Text = finalName;
        FinalNameBorder.Visibility = Visibility.Visible;
    }

    private void OpenWordDocument(string filePath)
    {
        try
        {
            _logger.Log($"Apertura documento in Word: {filePath}");

            
            var startInfo = new ProcessStartInfo
            {
                FileName = filePath,
                UseShellExecute = true
            };
            Process.Start(startInfo);

            
            if (_currentEditorState != null)
            {
                _currentEditorState.IsDocumentOpened = true;
            }

            
            StartFileWatcher(filePath);

            
            StartWordCheckTimer();

            _isWordOpen = true;
            UpdateWordOpenIndicator();

            _logger.Log("Word avviato - in attesa di modifiche e chiusura");
        }
        catch (Exception ex)
        {
            _logger.LogError("Errore durante l'apertura di Word", ex);
            MessageBox.Show(
                $"Impossibile aprire Word:\n\n{ex.Message}",
                "Errore Word",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    
    
    
    private void StartFileWatcher(string filePath)
    {
        StopFileWatcher();

        try
        {
            var directory = Path.GetDirectoryName(filePath);
            var fileName = Path.GetFileName(filePath);

            if (string.IsNullOrEmpty(directory)) return;

            _fileWatcher = new FileSystemWatcher(directory, fileName)
            {
                NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.FileName,
                EnableRaisingEvents = true
            };

            _fileWatcher.Changed += OnFileChanged;
            _fileWatcher.Created += OnFileChanged;
            _fileWatcher.Renamed += OnFileChanged;

            _logger.Log($"FileWatcher avviato per: {fileName}");
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Impossibile avviare FileWatcher: {ex.Message}");
        }
    }

    private void StopFileWatcher()
    {
        if (_fileWatcher != null)
        {
            _fileWatcher.EnableRaisingEvents = false;
            _fileWatcher.Changed -= OnFileChanged;
            _fileWatcher.Created -= OnFileChanged;
            _fileWatcher.Renamed -= OnFileChanged;
            _fileWatcher.Dispose();
            _fileWatcher = null;
        }
    }

    private void OnFileChanged(object sender, FileSystemEventArgs e)
    {
        
        Dispatcher.InvokeAsync(() =>
        {
            try
            {
                if (_currentEditorState == null) return;

                
                if (!File.Exists(_currentEditorState.LocalFilePath))
                {
                    _logger.LogWarning("File temporaneo non trovato durante il controllo modifica.");
                    return;
                }

                var currentModTime = File.GetLastWriteTime(_currentEditorState.LocalFilePath);
                if (currentModTime > _lastFileModifiedTime)
                {
                    _lastFileModifiedTime = currentModTime;
                    _currentEditorState.IsDocumentSaved = true;
                    _currentEditorState.LastModified = currentModTime;

                    _logger.Log("Documento salvato in Word - rilevata modifica file");
                    UpdateConvertButtonState();
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Errore durante il controllo modifica file: {ex.Message}");
            }
        });
    }

    
    
    
    private void StartWordCheckTimer()
    {
        StopWordCheckTimer();

        _wordCheckTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(2)
        };
        _wordCheckTimer.Tick += OnWordCheckTimerTick;
        _wordCheckTimer.Start();
    }

    private void StopWordCheckTimer()
    {
        if (_wordCheckTimer != null)
        {
            _wordCheckTimer.Stop();
            _wordCheckTimer.Tick -= OnWordCheckTimerTick;
            _wordCheckTimer = null;
        }
    }

    private void OnWordCheckTimerTick(object? sender, EventArgs e)
    {
        if (_currentEditorState == null) return;

        
        var isWordStillOpen = IsFileLockedByWord(_currentEditorState.LocalFilePath);

        if (_isWordOpen && !isWordStillOpen)
        {
            
            _isWordOpen = false;
            _logger.Log("Word chiuso - documento non più in editing");

            StopWordCheckTimer();
            StopFileWatcher();

            
            CheckFinalFileState();

            UpdateWordOpenIndicator();
            UpdateConvertButtonState();
        }
    }

    
    
    
    private bool IsFileLockedByWord(string filePath)
    {
        try
        {
            if (!File.Exists(filePath))
            {
                return false;
            }

            
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            return false; 
        }
        catch (IOException)
        {
            return true;
        }
        catch (UnauthorizedAccessException)
        {
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Errore verifica lock file: {ex.Message}");
            return false;
        }
    }

    
    
    
    private void CheckFinalFileState()
    {
        if (_currentEditorState == null) return;

        try
        {
            if (!File.Exists(_currentEditorState.LocalFilePath))
            {
                _logger.LogWarning("File temporaneo non trovato durante verifica finale.");
                return;
            }

            var currentModTime = File.GetLastWriteTime(_currentEditorState.LocalFilePath);
            if (currentModTime > _lastFileModifiedTime)
            {
                _lastFileModifiedTime = currentModTime;
                _currentEditorState.IsDocumentSaved = true;
                _currentEditorState.LastModified = currentModTime;
                _logger.Log("Verifica finale: documento risulta salvato");
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Errore durante verifica finale: {ex.Message}");
        }
    }

    private void UpdateWordOpenIndicator()
    {
        WordOpenIndicator.Visibility = _isWordOpen ? Visibility.Visible : Visibility.Collapsed;
    }

    
    
    
    private void EditSignatureButton_Click(object sender, RoutedEventArgs e)
    {
        if (_currentEditorState == null || string.IsNullOrEmpty(_selectedFilePath))
        {
            MessageBox.Show(
                "Nessuna firma selezionata da modificare.",
                "Attenzione",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        
        if (!_wordEditorService.ValidateEditorState(_currentEditorState))
        {
            MessageBox.Show(
                "Il file temporaneo non esiste più. Seleziona nuovamente la firma.",
                "File non trovato",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);

            _currentEditorState = null;
            UpdateConvertButtonState();
            return;
        }

        
        OpenWordDocument(_currentEditorState.LocalFilePath);
    }



    private void PresetListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (PresetListBox.SelectedItem is not PresetFile preset) return;

        LoadDocumentForEditing(
            preset.FullPath,
            preset.FileName,
            preset.DisplayName,
            "Preset");
    }

    private void LoadCustomButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "Seleziona documento",
            Filter = "Documenti supportati (*.docx;*.doc;*.rtf)|*.docx;*.doc;*.rtf|Documenti Word (*.docx;*.doc)|*.docx;*.doc|File RTF (*.rtf)|*.rtf",
            CheckFileExists = true
        };

        if (dialog.ShowDialog() == true)
        {
            PresetListBox.SelectedItem = null;
            var fileName = Path.GetFileName(dialog.FileName);
            var proposedName = Path.GetFileNameWithoutExtension(fileName);

            LoadDocumentForEditing(
                dialog.FileName,
                fileName,
                proposedName,
                "File personalizzato");
        }
    }

    private void LoadDocumentForEditing(string sourceFilePath, string displayFileName, string proposedSignatureName, string sourceLabel)
    {
        try
        {
            if (!File.Exists(sourceFilePath))
            {
                MessageBox.Show(
                    "Il file selezionato non esiste più. Riprova.",
                    "File non trovato",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            if (!IsSupportedDocument(sourceFilePath))
            {
                MessageBox.Show(
                    "Il file selezionato non è un documento supportato (.doc/.docx/.rtf).",
                    "File non valido",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            var normalizedName = WordConversionService.SanitizeFileName(proposedSignatureName);
            if (!string.Equals(proposedSignatureName, normalizedName, StringComparison.Ordinal))
            {
                _logger.LogWarning($"Nome firma normalizzato in import: '{proposedSignatureName}' → '{normalizedName}'");
            }

            var editableSource = PrepareSourceForEditing(sourceFilePath);

            _currentEditorState = _wordEditorService.PrepareFileForEditing(editableSource, normalizedName);
            _selectedFilePath = _currentEditorState.LocalFilePath;
            _lastFileModifiedTime = File.GetLastWriteTime(_currentEditorState.LocalFilePath);

            SelectedFileText.Text = displayFileName;
            SignatureNameTextBox.Text = normalizedName;

            _logger.Log($"{sourceLabel} importato: {displayFileName}");

            UpdateConvertButtonState();
        }
        catch (Exception ex)
        {
            _logger.LogError($"Errore durante l'import del documento", ex);
            MessageBox.Show(
                $"Errore durante l'import del documento:\n{ex.Message}",
                "Errore",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private string PrepareSourceForEditing(string sourceFilePath)
    {
        if (TempFileManager.IsUncPath(sourceFilePath))
        {
            _logger.Log("File su rete: copia in temporanea locale.");
            return _tempFileManager.CopyToLocalTemp(sourceFilePath);
        }

        return sourceFilePath;
    }

    private static bool IsSupportedDocument(string filePath)
    {
        var extension = Path.GetExtension(filePath);
        return extension.Equals(".doc", StringComparison.OrdinalIgnoreCase)
            || extension.Equals(".docx", StringComparison.OrdinalIgnoreCase)
            || extension.Equals(".rtf", StringComparison.OrdinalIgnoreCase);
    }

    private void SignatureNameTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        UpdateConvertButtonState();
    }

    private void AccountComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        UpdateFinalSignatureName();
    }

    private void BrowseFolderButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new System.Windows.Forms.FolderBrowserDialog
        {
            Description = "Seleziona cartella di destinazione per le firme",
            ShowNewFolderButton = true
        };

        if (!string.IsNullOrEmpty(DestinationFolderTextBox.Text) &&
            Directory.Exists(DestinationFolderTextBox.Text))
        {
            dialog.SelectedPath = DestinationFolderTextBox.Text;
        }

        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            DestinationFolderTextBox.Text = dialog.SelectedPath;
            ValidateDestinationFolder(dialog.SelectedPath);
            RefreshExistingSignatures();
        }
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

        
        string? identifier = null;
        if (_isOutlookAvailable && AccountComboBox.SelectedItem is OutlookAccount account)
        {
            identifier = account.DisplayText;
        }
        else if (!string.IsNullOrWhiteSpace(IdentifierTextBox.Text))
        {
            identifier = IdentifierTextBox.Text.Trim();
        }

        var finalSignatureName = _signatureWorkflowService.BuildFinalSignatureName(baseName, identifier);
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
            var fixOutlook2512 = FixOutlook2512CheckBox.IsChecked == true;

            var conversionResult = await Task.Run(() =>
            {
                return _signatureWorkflowService.ConvertDocument(
                    _selectedFilePath,
                    destinationFolder,
                    finalSignatureName,
                    useFilteredHtml,
                    fixOutlook2512);
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

    private void ExistingSignaturesListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        DeleteSignatureButton.IsEnabled = ExistingSignaturesListBox.SelectedItem != null;
    }

    private void BackupsListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        UpdateBackupButtons();
    }

    private void RestoreBackupButton_Click(object sender, RoutedEventArgs e)
    {
        if (BackupsListBox.SelectedItem is not BackupInfo backup)
        {
            return;
        }

        var result = MessageBox.Show(
            $"Ripristinare il backup '{backup.FileName}'?\n\n" +
            "I file nella cartella firme verranno sovrascritti.",
            "Conferma ripristino",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (result != MessageBoxResult.Yes)
        {
            return;
        }

        SetBusy(true, "Ripristino backup in corso...");

        try
        {
            var destinationFolder = SignatureRepository.GetDefaultOutlookSignaturesFolder();
            var restored = _signatureRepository.RestoreBackup(backup, destinationFolder);

            if (restored)
            {
                _logger.Log($"Backup ripristinato: {backup.FileName}");
                MessageBox.Show("Backup ripristinato con successo.", "Ripristino completato", MessageBoxButton.OK,
                    MessageBoxImage.Information);
                RefreshExistingSignatures();
            }
            else
            {
                MessageBox.Show("Impossibile ripristinare il backup selezionato.", "Errore ripristino",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        finally
        {
            SetBusy(false);
        }
    }

    private void DeleteBackupButton_Click(object sender, RoutedEventArgs e)
    {
        if (BackupsListBox.SelectedItem is not BackupInfo backup)
        {
            return;
        }

        var result = MessageBox.Show(
            $"Eliminare il backup '{backup.FileName}'?",
            "Conferma eliminazione",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);

        if (result != MessageBoxResult.Yes)
        {
            return;
        }

        try
        {
            if (_signatureRepository.DeleteBackup(backup))
            {
                _logger.Log($"Backup eliminato: {backup.FileName}");
            }
            else
            {
                MessageBox.Show("Impossibile eliminare il backup selezionato.", "Errore eliminazione",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        finally
        {
            RefreshBackups();
        }
    }

    private void RefreshBackupsButton_Click(object sender, RoutedEventArgs e)
    {
        RefreshBackups();
    }

    private void DeleteSignatureButton_Click(object sender, RoutedEventArgs e)
    {
        if (ExistingSignaturesListBox.SelectedItem is not SignatureInfo signature)
        {
            return;
        }

        var result = MessageBox.Show(
            $"Eliminare la firma '{signature.Name}'?\n\n" +
            "Verranno eliminati tutti i file associati (.htm, .rtf, .txt e cartelle assets).",
            "Conferma eliminazione",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);

        if (result != MessageBoxResult.Yes)
        {
            return;
        }

        try
        {
            _signatureRepository.DeleteSignature(signature);
            RefreshExistingSignatures();
            _logger.Log($"Firma '{signature.Name}' eliminata");
            CheckOverwriteWarning();
        }
        catch (Exception ex)
        {
            _logger.LogError($"Errore durante l'eliminazione della firma", ex);
            MessageBox.Show(
                $"Errore durante l'eliminazione:\n{ex.Message}",
                "Errore",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private void RefreshSignaturesButton_Click(object sender, RoutedEventArgs e)
    {
        RefreshExistingSignatures();
    }

    private void BrowseSignaturesButton_Click(object sender, RoutedEventArgs e)
    {
        OpenDefaultSignaturesFolder();
    }

    private void BrowseBackupsButton_Click(object sender, RoutedEventArgs e)
    {
        OpenDefaultSignaturesFolder();
    }

    private void CopyLogButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            Clipboard.SetText(LogTextBox.Text);
            _logger.Log("Log copiato negli appunti");
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Impossibile copiare negli appunti: {ex.Message}");
        }
    }

    private void ClearLogButton_Click(object sender, RoutedEventArgs e)
    {
        _logger.Clear();
        LogTextBox.Clear();
        _logger.Log("Log pulito");
    }

    private void OpenLogFileButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var logFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OnlyFirmaOutlook",
                "Logs");

            var logFile = Path.Combine(logFolder, "app.log");

            if (File.Exists(logFile))
            {
                Process.Start("notepad.exe", logFile);
            }
            else
            {
                Process.Start("explorer.exe", logFolder);
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Impossibile aprire il file di log: {ex.Message}");
        }
    }

    private void OpenDefaultSignaturesFolder()
    {
        try
        {
            var defaultFolder = SignatureRepository.GetDefaultOutlookSignaturesFolder();
            if (Directory.Exists(defaultFolder))
            {
                Process.Start("explorer.exe", defaultFolder);
            }
            else
            {
                MessageBox.Show("La cartella firme predefinita non è disponibile.", "Cartella non trovata",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Impossibile aprire la cartella firme: {ex.Message}");
        }
    }



    private void SetBusy(bool isBusy, string? message = null)
    {
        BusyOverlay.Visibility = isBusy ? Visibility.Visible : Visibility.Collapsed;

        if (!string.IsNullOrEmpty(message))
        {
            BusyMessage.Text = message;
        }

        
        PresetListBox.IsEnabled = !isBusy;
        LoadCustomButton.IsEnabled = !isBusy;
        SignatureNameTextBox.IsEnabled = !isBusy;
        AccountComboBox.IsEnabled = !isBusy;
        IdentifierTextBox.IsEnabled = !isBusy;
        BrowseFolderButton.IsEnabled = !isBusy;
        FilteredHtmlRadio.IsEnabled = !isBusy;
        CompleteHtmlRadio.IsEnabled = !isBusy;
        FixOutlook2512CheckBox.IsEnabled = !isBusy;
        if (isBusy)
        {
            ConvertButton.IsEnabled = false;
        }
        else
        {
            UpdateConvertButtonState();
        }
        DeleteSignatureButton.IsEnabled = !isBusy && ExistingSignaturesListBox.SelectedItem != null;
        RefreshSignaturesButton.IsEnabled = !isBusy;
        BrowseSignaturesButton.IsEnabled = !isBusy;
        RestoreBackupButton.IsEnabled = !isBusy && BackupsListBox.SelectedItem != null;
        DeleteBackupButton.IsEnabled = !isBusy && BackupsListBox.SelectedItem != null;
        RefreshBackupsButton.IsEnabled = !isBusy;
        BrowseBackupsButton.IsEnabled = !isBusy;
    }

    private void OnLogAdded(object? sender, string message)
    {
        Dispatcher.InvokeAsync(() =>
        {
            LogTextBox.AppendText(message + Environment.NewLine);
            ScrollLogToEnd();
        });
    }

    private void ScrollLogToEnd()
    {
        LogTextBox.ScrollToEnd();
    }

    private void ResetUiForNewSignature()
    {
        _selectedFilePath = null;
        _currentEditorState = null;
        _isWordOpen = false;

        PresetListBox.SelectedItem = null;
        SelectedFileText.Text = "Nessun file selezionato";
        SignatureNameTextBox.Text = string.Empty;
        IdentifierTextBox.Text = string.Empty;
        AccountComboBox.SelectedIndex = _accounts.Count > 0 ? 0 : -1;
        FilteredHtmlRadio.IsChecked = false;
        CompleteHtmlRadio.IsChecked = true;
        ExistingSignaturesListBox.SelectedItem = null;
        BackupsListBox.SelectedItem = null;

        UpdateWordOpenIndicator();
        RefreshExistingSignatures();
        RefreshBackups();
        UpdateConvertButtonState();
    }

    private void GuideToggleButton_Checked(object sender, RoutedEventArgs e)
    {
        if (_guideWindow == null)
        {
            _guideWindow = new GuideWindow
            {
                Owner = this
            };
            _guideWindow.Closed += GuideWindow_Closed;
        }

        _guideWindow.Show();
        _guideWindow.Activate();
    }

    private void GuideToggleButton_Unchecked(object sender, RoutedEventArgs e)
    {
        if (_guideWindow != null)
        {
            _guideWindow.Closed -= GuideWindow_Closed;
            _guideWindow.Close();
            _guideWindow = null;
        }
    }

    private void GuideWindow_Closed(object? sender, EventArgs e)
    {
        if (_guideWindow != null)
        {
            _guideWindow.Closed -= GuideWindow_Closed;
            _guideWindow = null;
        }

        GuideToggleButton.IsChecked = false;
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
