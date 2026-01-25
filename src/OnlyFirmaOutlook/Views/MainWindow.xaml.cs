using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using OnlyFirmaOutlook.Models;
using OnlyFirmaOutlook.Services;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using MessageBox = System.Windows.MessageBox;
using Clipboard = System.Windows.Clipboard;

namespace OnlyFirmaOutlook.Views;

/// <summary>
/// MainWindow - Finestra principale dell'applicazione.
/// Gestisce l'interfaccia utente per la conversione di documenti Word in firme Outlook.
/// </summary>
public partial class MainWindow : Window
{
    private readonly LoggingService _logger;
    private readonly TempFileManager _tempFileManager;
    private readonly PresetService _presetService;
    private readonly OutlookAccountService _outlookAccountService;
    private readonly SignatureRepository _signatureRepository;
    private readonly WordConversionService _wordConversionService;
    private readonly WordEditorService _wordEditorService;

    private List<PresetFile> _presets = new();
    private List<OutlookAccount> _accounts = new();
    private List<SignatureInfo> _existingSignatures = new();

    private string? _selectedFilePath;
    private EditorState? _currentEditorState;
    private bool _isOutlookAvailable;
    private bool _isFolderWritable;

    // Word editing state
    private FileSystemWatcher? _fileWatcher;
    private DispatcherTimer? _wordCheckTimer;
    private DateTime _lastFileModifiedTime;
    private bool _isWordOpen;

    public MainWindow()
    {
        InitializeComponent();

        // Inizializza servizi
        _logger = LoggingService.Instance;
        _tempFileManager = TempFileManager.Instance;
        _presetService = new PresetService();
        _outlookAccountService = new OutlookAccountService();
        _signatureRepository = new SignatureRepository();
        _wordConversionService = new WordConversionService();
        _wordEditorService = new WordEditorService();

        // Sottoscrivi agli eventi di log
        _logger.LogAdded += OnLogAdded;

        // Carica il log esistente
        LogTextBox.Text = _logger.GetFullLog();
        ScrollLogToEnd();

        // Inizializza l'applicazione
        Loaded += MainWindow_Loaded;
    }

    private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        _logger.Log("Inizializzazione interfaccia...");

        // Pulisci cartelle temporanee orfane
        _tempFileManager.CleanupOrphanedFolders();
        _wordEditorService.CleanupOrphanedEditorFolders();

        // Carica preset
        LoadPresets();

        // Inizializza in modo asincrono
        await InitializeAsync();

        _logger.Log("Interfaccia pronta");
    }

    private async Task InitializeAsync()
    {
        SetBusy(true, "Rilevamento configurazione Office...");

        try
        {
            // Verifica Word
            if (!OfficeBitnessDetector.IsWordInstalled())
            {
                MessageBox.Show(
                    "Microsoft Word non risulta installato.\n\n" +
                    "Word è necessario per la conversione dei documenti.",
                    "Word non trovato",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }

            // Carica account Outlook
            var accountResult = await Task.Run(() => _outlookAccountService.LoadAccounts());

            _isOutlookAvailable = accountResult.OutlookAvailable;
            _accounts = accountResult.Accounts;

            // Configura UI in base a disponibilità Outlook
            ConfigureOutlookUI();

            // Imposta cartella destinazione predefinita
            SetDefaultDestinationFolder();

            // Carica firme esistenti
            RefreshExistingSignatures();
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
            // Outlook disponibile con account
            OutlookWarningBorder.Visibility = Visibility.Collapsed;
            AccountLabel.Visibility = Visibility.Visible;
            AccountComboBox.Visibility = Visibility.Visible;
            AccountComboBox.ItemsSource = _accounts;
            AccountComboBox.DisplayMemberPath = "DisplayText";

            IdentifierLabel.Visibility = Visibility.Collapsed;
            IdentifierTextBox.Visibility = Visibility.Collapsed;
            IdentifierHint.Visibility = Visibility.Collapsed;

            // Seleziona il primo account
            if (_accounts.Count > 0)
            {
                AccountComboBox.SelectedIndex = 0;
            }
        }
        else if (_isOutlookAvailable && _accounts.Count == 0)
        {
            // Outlook disponibile ma senza account configurati
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
            // Outlook non disponibile
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

    private void UpdateConvertButtonState()
    {
        var hasFile = !string.IsNullOrEmpty(_selectedFilePath) && File.Exists(_selectedFilePath);
        var hasSignatureName = !string.IsNullOrWhiteSpace(SignatureNameTextBox.Text);
        var hasDestination = !string.IsNullOrWhiteSpace(DestinationFolderTextBox.Text);
        var isDocumentReady = _currentEditorState?.IsReadyForConversion ?? false;

        // Il pulsante Converti è abilitato SOLO se il documento è stato modificato e salvato
        ConvertButton.IsEnabled = hasFile && hasSignatureName && hasDestination && _isFolderWritable && isDocumentReady;

        // Aggiorna stato modifica firma
        UpdateEditStatusDisplay();

        UpdateFinalSignatureName();

        // Aggiorna evidenziazione step
        UpdateStepHighlighting();

        // Verifica sovrascrittura firme esistenti
        CheckOverwriteWarning();
    }

    /// <summary>
    /// Aggiorna l'evidenziazione degli step in base allo stato corrente.
    /// </summary>
    private void UpdateStepHighlighting()
    {
        var hasSignatureSelected = _currentEditorState != null;
        var hasSignatureName = !string.IsNullOrWhiteSpace(SignatureNameTextBox.Text);
        var hasDestination = _isFolderWritable;
        var isDocumentReady = _currentEditorState?.IsReadyForConversion ?? false;

        // Step 1: Selezione firma
        if (!hasSignatureSelected)
        {
            SetStepStyle(Step1Group, StepState.Current);
            SetStepStyle(Step2Group, StepState.Pending);
            SetStepStyle(Step3Group, StepState.Pending);
            SetStepStyle(Step4Group, StepState.Pending);
            SetStepStyle(Step5Group, StepState.Pending);
            SetStepStyle(Step6Group, StepState.Pending);
            return;
        }

        SetStepStyle(Step1Group, StepState.Completed);

        // Step 2: Nome e account
        if (!hasSignatureName)
        {
            SetStepStyle(Step2Group, StepState.Current);
            SetStepStyle(Step3Group, StepState.Pending);
            SetStepStyle(Step4Group, StepState.Pending);
            SetStepStyle(Step5Group, StepState.Pending);
            SetStepStyle(Step6Group, StepState.Pending);
            return;
        }

        SetStepStyle(Step2Group, StepState.Completed);

        // Step 3: Cartella destinazione
        if (!hasDestination)
        {
            SetStepStyle(Step3Group, StepState.Current);
            SetStepStyle(Step4Group, StepState.Pending);
            SetStepStyle(Step5Group, StepState.Pending);
            SetStepStyle(Step6Group, StepState.Pending);
            return;
        }

        SetStepStyle(Step3Group, StepState.Completed);

        // Step 4: Modifica firma
        if (!isDocumentReady)
        {
            SetStepStyle(Step4Group, StepState.Current);
            SetStepStyle(Step5Group, StepState.Pending);
            SetStepStyle(Step6Group, StepState.Pending);
            return;
        }

        SetStepStyle(Step4Group, StepState.Completed);

        // Step 5 e 6: Completati (opzionali)
        SetStepStyle(Step5Group, StepState.Completed);
        SetStepStyle(Step6Group, StepState.Completed);
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

    /// <summary>
    /// Verifica se la firma corrente sovrascriverà una esistente.
    /// </summary>
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

        var finalName = WordConversionService.GenerateSignatureName(baseName, identifier);
        var destinationFolder = DestinationFolderTextBox.Text;

        if (!string.IsNullOrEmpty(destinationFolder) && _signatureRepository.SignatureExists(destinationFolder, finalName))
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

        var finalName = WordConversionService.GenerateSignatureName(baseName, identifier);

        FinalSignatureNameText.Text = finalName;
        FinalNameBorder.Visibility = Visibility.Visible;
    }

    #region Word Editor Methods

    /// <summary>
    /// Prepara un file per l'editing e apre Word direttamente.
    /// </summary>
    private void PrepareAndOpenInWord(string sourceFilePath, string proposedSignatureName)
    {
        try
        {
            _logger.Log($"Preparazione file per Word: {proposedSignatureName}");

            // Prepara il file per l'editing (copia in cartella EditorTemp dedicata)
            _currentEditorState = _wordEditorService.PrepareFileForEditing(sourceFilePath, proposedSignatureName);

            // Aggiorna il percorso del file selezionato con la copia nell'EditorTemp
            _selectedFilePath = _currentEditorState.LocalFilePath;

            // Salva il timestamp iniziale del file
            _lastFileModifiedTime = File.GetLastWriteTime(_currentEditorState.LocalFilePath);

            // Apri Word direttamente
            OpenWordDocument(_currentEditorState.LocalFilePath);

            UpdateConvertButtonState();
        }
        catch (Exception ex)
        {
            _logger.LogError("Errore durante la preparazione del file", ex);
            MessageBox.Show(
                $"Errore durante la preparazione del file:\n\n{ex.Message}",
                "Errore",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    /// <summary>
    /// Apre un documento in Word e avvia il monitoraggio.
    /// </summary>
    private void OpenWordDocument(string filePath)
    {
        try
        {
            _logger.Log($"Apertura documento in Word: {filePath}");

            // Avvia Word con il documento
            var startInfo = new ProcessStartInfo
            {
                FileName = filePath,
                UseShellExecute = true
            };
            Process.Start(startInfo);

            // Marca come aperto
            if (_currentEditorState != null)
            {
                _currentEditorState.IsDocumentOpened = true;
            }

            // Avvia monitoraggio file
            StartFileWatcher(filePath);

            // Avvia timer per verificare se Word è ancora aperto
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

    /// <summary>
    /// Avvia il FileSystemWatcher per monitorare le modifiche al file.
    /// </summary>
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
                NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size,
                EnableRaisingEvents = true
            };

            _fileWatcher.Changed += OnFileChanged;

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
            _fileWatcher.Dispose();
            _fileWatcher = null;
        }
    }

    private void OnFileChanged(object sender, FileSystemEventArgs e)
    {
        // Esegui sul thread UI
        Dispatcher.InvokeAsync(() =>
        {
            try
            {
                if (_currentEditorState == null) return;

                // Verifica se il file è stato effettivamente modificato
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

    /// <summary>
    /// Avvia il timer per verificare periodicamente se Word è ancora aperto.
    /// </summary>
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

        // Verifica se ci sono processi Word che hanno il file aperto
        var isWordStillOpen = IsFileLockedByWord(_currentEditorState.LocalFilePath);

        if (_isWordOpen && !isWordStillOpen)
        {
            // Word è stato chiuso
            _isWordOpen = false;
            _logger.Log("Word chiuso - documento non più in editing");

            StopWordCheckTimer();
            StopFileWatcher();

            // Verifica finale se il file è stato modificato
            CheckFinalFileState();

            UpdateWordOpenIndicator();
            UpdateConvertButtonState();
        }
    }

    /// <summary>
    /// Verifica se il file è bloccato da Word.
    /// </summary>
    private bool IsFileLockedByWord(string filePath)
    {
        try
        {
            // Prova ad aprire il file in modo esclusivo
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            return false; // File non è bloccato
        }
        catch (IOException)
        {
            return true; // File è bloccato (probabilmente da Word)
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Verifica finale dello stato del file dopo la chiusura di Word.
    /// </summary>
    private void CheckFinalFileState()
    {
        if (_currentEditorState == null) return;

        try
        {
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

    /// <summary>
    /// Apre/riapre Word per il file corrente.
    /// </summary>
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

        // Verifica che il file locale esista ancora
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

        // Apri Word direttamente
        OpenWordDocument(_currentEditorState.LocalFilePath);
    }

    #endregion

    #region Event Handlers

    private void PresetListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (PresetListBox.SelectedItem is not PresetFile preset) return;

        try
        {
            // Copia il file dalla share (se necessario) alla cartella temporanea locale
            var tempFilePath = _tempFileManager.CopyToLocalTemp(preset.FullPath);

            // Prepara il file per l'editing (crea EditorState) ma NON apre Word
            _currentEditorState = _wordEditorService.PrepareFileForEditing(tempFilePath, preset.DisplayName);
            _selectedFilePath = _currentEditorState.LocalFilePath;
            _lastFileModifiedTime = File.GetLastWriteTime(_currentEditorState.LocalFilePath);

            // Aggiorna UI
            SelectedFileText.Text = preset.FileName;
            SignatureNameTextBox.Text = preset.DisplayName;

            _logger.Log($"Preset selezionato: {preset.DisplayName}");

            UpdateConvertButtonState();
        }
        catch (Exception ex)
        {
            _logger.LogError($"Errore durante la selezione del preset", ex);
            MessageBox.Show(
                $"Errore durante la selezione del preset:\n{ex.Message}",
                "Errore",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private void LoadCustomButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "Seleziona documento Word",
            Filter = "Documenti Word (*.docx;*.doc)|*.docx;*.doc|Tutti i file (*.*)|*.*",
            CheckFileExists = true
        };

        if (dialog.ShowDialog() == true)
        {
            try
            {
                // Deseleziona eventuale preset
                PresetListBox.SelectedItem = null;

                // Se il file è su una share di rete, copialo localmente
                string sourceFile;
                if (TempFileManager.IsUncPath(dialog.FileName))
                {
                    sourceFile = _tempFileManager.CopyToLocalTemp(dialog.FileName);
                }
                else
                {
                    sourceFile = dialog.FileName;
                }

                var fileName = Path.GetFileName(dialog.FileName);
                var proposedName = Path.GetFileNameWithoutExtension(fileName);

                // Prepara il file per l'editing (crea EditorState) ma NON apre Word
                _currentEditorState = _wordEditorService.PrepareFileForEditing(sourceFile, proposedName);
                _selectedFilePath = _currentEditorState.LocalFilePath;
                _lastFileModifiedTime = File.GetLastWriteTime(_currentEditorState.LocalFilePath);

                // Aggiorna UI
                SelectedFileText.Text = fileName;
                SignatureNameTextBox.Text = proposedName;

                _logger.Log($"File personalizzato caricato: {fileName}");

                UpdateConvertButtonState();
            }
            catch (Exception ex)
            {
                _logger.LogError($"Errore durante il caricamento del file", ex);
                MessageBox.Show(
                    $"Errore durante il caricamento del file:\n{ex.Message}",
                    "Errore",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }
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

        // Verifica che il documento sia stato modificato e salvato
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

        // Determina l'identificativo
        string? identifier = null;
        if (_isOutlookAvailable && AccountComboBox.SelectedItem is OutlookAccount account)
        {
            identifier = account.DisplayText;
        }
        else if (!string.IsNullOrWhiteSpace(IdentifierTextBox.Text))
        {
            identifier = IdentifierTextBox.Text.Trim();
        }

        var finalSignatureName = WordConversionService.GenerateSignatureName(baseName, identifier);
        var destinationFolder = DestinationFolderTextBox.Text;

        // Verifica se esiste già una firma con questo nome
        if (_signatureRepository.SignatureExists(destinationFolder, finalSignatureName))
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

            // Elimina i file esistenti
            _signatureRepository.DeleteExistingSignatureFiles(destinationFolder, finalSignatureName);
        }

        SetBusy(true, "Conversione in corso...");

        try
        {
            var useFilteredHtml = FilteredHtmlRadio.IsChecked == true;

            // Esegui la conversione su un thread in background (ma con STA)
            var conversionResult = await Task.Run(() =>
            {
                return _wordConversionService.ConvertDocument(
                    _selectedFilePath,
                    destinationFolder,
                    finalSignatureName,
                    useFilteredHtml);
            });

            if (conversionResult.Success)
            {
                _logger.Log("Conversione completata con successo!");

                // Cleanup cartella EditorTemp dopo conversione riuscita
                if (_currentEditorState != null)
                {
                    _wordEditorService.CleanupEditorTempFolder(_currentEditorState.EditorSessionId);
                    _currentEditorState = null;
                }

                // Reset stato
                _selectedFilePath = null;
                UpdateConvertButtonState();

                // Aggiorna lista firme
                RefreshExistingSignatures();

                // Apri Esplora File nella cartella di destinazione
                OpenDestinationFolder(destinationFolder, conversionResult.HtmFilePath);

                MessageBox.Show(
                    $"Firma '{finalSignatureName}' creata con successo!\n\n" +
                    $"File creati:\n" +
                    $"- {Path.GetFileName(conversionResult.HtmFilePath)}\n" +
                    $"- {Path.GetFileName(conversionResult.RtfFilePath)}\n" +
                    $"- {Path.GetFileName(conversionResult.TxtFilePath)}" +
                    (conversionResult.AssetsFolderPath != null
                        ? $"\n- {Path.GetFileName(conversionResult.AssetsFolderPath)}/"
                        : ""),
                    "Conversione completata",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
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

    private void OpenDestinationFolder(string folderPath, string? htmFilePath)
    {
        try
        {
            if (!string.IsNullOrEmpty(htmFilePath) && File.Exists(htmFilePath))
            {
                // Apri Esplora File e seleziona il file HTM
                Process.Start("explorer.exe", $"/select,\"{htmFilePath}\"");
            }
            else if (Directory.Exists(folderPath))
            {
                // Apri solo la cartella
                Process.Start("explorer.exe", folderPath);
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Impossibile aprire Esplora File: {ex.Message}");
        }
    }

    private void ExistingSignaturesListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        DeleteSignatureButton.IsEnabled = ExistingSignaturesListBox.SelectedItem != null;
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

    #endregion

    #region UI Helpers

    private void SetBusy(bool isBusy, string? message = null)
    {
        BusyOverlay.Visibility = isBusy ? Visibility.Visible : Visibility.Collapsed;

        if (!string.IsNullOrEmpty(message))
        {
            BusyMessage.Text = message;
        }

        // Disabilita/abilita i controlli
        PresetListBox.IsEnabled = !isBusy;
        LoadCustomButton.IsEnabled = !isBusy;
        SignatureNameTextBox.IsEnabled = !isBusy;
        AccountComboBox.IsEnabled = !isBusy;
        IdentifierTextBox.IsEnabled = !isBusy;
        BrowseFolderButton.IsEnabled = !isBusy;
        FilteredHtmlRadio.IsEnabled = !isBusy;
        CompleteHtmlRadio.IsEnabled = !isBusy;
        ConvertButton.IsEnabled = !isBusy && _isFolderWritable;
        DeleteSignatureButton.IsEnabled = !isBusy && ExistingSignaturesListBox.SelectedItem != null;
        RefreshSignaturesButton.IsEnabled = !isBusy;
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

    #endregion

    protected override void OnClosed(EventArgs e)
    {
        // Cleanup watcher e timer
        StopFileWatcher();
        StopWordCheckTimer();

        _logger.LogAdded -= OnLogAdded;
        base.OnClosed(e);
    }
}
