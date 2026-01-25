using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Threading;
using OnlyFirmaOutlook.Helpers;
using OnlyFirmaOutlook.Models;
using OnlyFirmaOutlook.Services;
using OnlyFirmaOutlook.ViewModels;
using MessageBox = System.Windows.MessageBox;

namespace OnlyFirmaOutlook.Windows;

/// <summary>
/// Finestra editor Word integrato.
/// Gestisce l'apertura, modifica e salvataggio di documenti Word tramite COM.
/// Thread STA dedicato per tutte le operazioni COM.
/// </summary>
public partial class WordEditorWindow : Window
{
    private readonly LoggingService _logger;
    private readonly WordEditorViewModel _viewModel;
    private readonly EditorState _editorState;

    private dynamic? _wordApp;
    private dynamic? _wordDocument;
    private bool _documentLoaded;
    private bool _closingConfirmed;

    /// <summary>
    /// Risultato della finestra: true se salvato e pronto per conversione.
    /// </summary>
    public bool DocumentReadyForConversion { get; private set; }

    public WordEditorWindow(EditorState editorState)
    {
        InitializeComponent();

        _logger = LoggingService.Instance;
        _editorState = editorState;
        _viewModel = new WordEditorViewModel(editorState);

        DataContext = _viewModel;
        Title = _viewModel.WindowTitle;

        _logger.Log($"Apertura editor per: {editorState.ProposedSignatureName}");

        Loaded += WordEditorWindow_Loaded;
    }

    private void BringWordToFrontButton_Click(object sender, RoutedEventArgs e)
    {
        BringWordToFront();
    }

    private void BringWordToFront()
    {
        try
        {
            if (_wordApp != null)
            {
                _wordApp.Activate();
                var activeWindow = _wordApp.ActiveWindow;
                if (activeWindow != null)
                {
                    var hwnd = new IntPtr((int)activeWindow.Hwnd);
                    NativeMethods.SetForegroundWindow(hwnd);
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Errore durante l'attivazione di Word: {ex.Message}");
        }
    }

    private async void WordEditorWindow_Loaded(object sender, RoutedEventArgs e)
    {
        await LoadWordDocumentAsync();
    }

    private async Task LoadWordDocumentAsync()
    {
        SetBusy(true, "Apertura documento Word...");

        try
        {
            // Verifica che il file esista
            if (!File.Exists(_editorState.LocalFilePath))
            {
                throw new FileNotFoundException($"File non trovato: {_editorState.LocalFilePath}");
            }

            // Carica Word direttamente sul thread UI (già STA)
            // IMPORTANTE: Non usare thread separato per evitare RCW separation
            LoadWordOnUiThread();

            // Piccolo delay per permettere a Word di inizializzare
            await Task.Delay(500);

            // Marca documento come aperto
            _viewModel.MarkDocumentOpened();
            _documentLoaded = true;

            UpdateUI();
            _logger.Log("Documento Word caricato nell'editor");
        }
        catch (Exception ex)
        {
            _logger.LogError("Errore durante il caricamento del documento Word", ex);

            MessageBox.Show(
                $"Impossibile aprire il documento Word:\n\n{ex.Message}\n\n" +
                "L'editor verrà chiuso.",
                "Errore apertura documento",
                MessageBoxButton.OK,
                MessageBoxImage.Error);

            _closingConfirmed = true;
            Close();
        }
        finally
        {
            SetBusy(false);
        }
    }

    /// <summary>
    /// Carica Word COM sul thread UI principale (già STA).
    /// Word viene aperto come finestra standalone per permettere piena interazione.
    /// </summary>
    private void LoadWordOnUiThread()
    {
        _logger.Log("Creazione istanza Word sul thread UI (STA)...");

        // Crea Word Application
        var wordType = Type.GetTypeFromProgID("Word.Application");
        if (wordType == null)
        {
            throw new COMException("Microsoft Word non è installato o non accessibile");
        }

        _wordApp = Activator.CreateInstance(wordType);
        if (_wordApp == null)
        {
            throw new COMException("Impossibile creare istanza di Word");
        }

        // Configurazione Word - visibile come finestra standalone
        _wordApp.Visible = true;
        _wordApp.DisplayAlerts = 0; // wdAlertsNone

        // Apri documento (non ReadOnly per permettere modifica)
        _logger.Log($"Apertura documento: {_editorState.LocalFilePath}");

        _wordDocument = _wordApp.Documents.Open(
            FileName: _editorState.LocalFilePath,
            ReadOnly: false,
            AddToRecentFiles: false,
            Visible: true);

        if (_wordDocument == null)
        {
            throw new COMException("Impossibile aprire il documento");
        }

        _logger.Log("Documento aperto con successo");

        // Porta Word in primo piano
        try
        {
            _wordApp.Activate();
            var wordWindowHandle = new IntPtr((int)_wordApp.ActiveWindow.Hwnd);
            NativeMethods.SetForegroundWindow(wordWindowHandle);
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Impossibile portare Word in primo piano: {ex.Message}");
        }
    }

    private void UpdateUI()
    {
        Dispatcher.Invoke(() =>
        {
            StatusTextBlock.Text = _viewModel.StatusText;

            UnsavedChangesTextBlock.Visibility = _viewModel.HasUnsavedChanges
                ? Visibility.Visible
                : Visibility.Collapsed;

            SaveButton.IsEnabled = _documentLoaded;
            SaveAndCloseButton.IsEnabled = _documentLoaded;
        });
    }

    private void SaveButton_Click(object sender, RoutedEventArgs e)
    {
        SaveDocument(false);
    }

    private void SaveAndCloseButton_Click(object sender, RoutedEventArgs e)
    {
        if (SaveDocument(true))
        {
            _closingConfirmed = true;
            DocumentReadyForConversion = true;
            Close();
        }
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    /// <summary>
    /// Salva il documento Word.
    /// </summary>
    /// <param name="silent">Se true, non mostra messaggi di successo</param>
    /// <returns>True se salvato con successo</returns>
    private bool SaveDocument(bool silent)
    {
        if (_wordDocument == null)
        {
            MessageBox.Show("Nessun documento aperto.", "Attenzione", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        SetBusy(true, "Salvataggio documento...");

        try
        {
            _logger.Log("Salvataggio documento Word...");
            _wordDocument.Save();

            _viewModel.MarkDocumentSaved();
            UpdateUI();

            _logger.Log("Documento salvato con successo");

            if (!silent)
            {
                MessageBox.Show(
                    "Documento salvato con successo!",
                    "Salvataggio completato",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }

            return true;
        }
        catch (COMException comEx)
        {
            _logger.LogError("Errore COM durante il salvataggio", comEx);

            MessageBox.Show(
                $"Errore durante il salvataggio del documento:\n\n{comEx.Message}",
                "Errore salvataggio",
                MessageBoxButton.OK,
                MessageBoxImage.Error);

            return false;
        }
        catch (Exception ex)
        {
            _logger.LogError("Errore durante il salvataggio", ex);

            MessageBox.Show(
                $"Errore durante il salvataggio:\n\n{ex.Message}",
                "Errore",
                MessageBoxButton.OK,
                MessageBoxImage.Error);

            return false;
        }
        finally
        {
            SetBusy(false);
        }
    }

    private void Window_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        // Se chiusura già confermata, procedi
        if (_closingConfirmed)
        {
            CleanupWord();
            return;
        }

        // Se ci sono modifiche non salvate, avvisa l'utente
        if (_viewModel.HasUnsavedChanges)
        {
            var result = MessageBox.Show(
                "Ci sono modifiche non salvate.\n\n" +
                "Vuoi salvare prima di chiudere?",
                "Modifiche non salvate",
                MessageBoxButton.YesNoCancel,
                MessageBoxImage.Question);

            switch (result)
            {
                case MessageBoxResult.Yes:
                    // Salva e chiudi
                    if (SaveDocument(true))
                    {
                        DocumentReadyForConversion = _editorState.IsReadyForConversion;
                        CleanupWord();
                    }
                    else
                    {
                        // Salvataggio fallito, annulla chiusura
                        e.Cancel = true;
                    }
                    break;

                case MessageBoxResult.No:
                    // Chiudi senza salvare
                    DocumentReadyForConversion = _editorState.IsReadyForConversion;
                    CleanupWord();
                    break;

                case MessageBoxResult.Cancel:
                    // Annulla chiusura
                    e.Cancel = true;
                    break;
            }
        }
        else
        {
            // Nessuna modifica, chiudi normalmente
            DocumentReadyForConversion = _editorState.IsReadyForConversion;
            CleanupWord();
        }
    }

    private void CleanupWord()
    {
        _logger.Log("Cleanup editor Word...");

        try
        {
            if (_wordDocument != null)
            {
                ComHelper.CloseWordDocument(_wordDocument, saveChanges: false);
                _wordDocument = null;
            }

            if (_wordApp != null)
            {
                ComHelper.QuitWordApplication(_wordApp, saveChanges: false);
                _wordApp = null;
            }

            ComHelper.ForceGarbageCollection();

            _logger.Log("Cleanup Word completato");
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Errore durante cleanup Word: {ex.Message}");
        }
    }

    private void SetBusy(bool isBusy, string? message = null)
    {
        Dispatcher.Invoke(() =>
        {
            BusyOverlay.Visibility = isBusy ? Visibility.Visible : Visibility.Collapsed;

            if (!string.IsNullOrEmpty(message))
            {
                BusyMessage.Text = message;
            }
        });
    }

    protected override void OnClosed(EventArgs e)
    {
        base.OnClosed(e);
    }
}

/// <summary>
/// Native methods per gestione finestra Word.
/// </summary>
internal static class NativeMethods
{
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
}
