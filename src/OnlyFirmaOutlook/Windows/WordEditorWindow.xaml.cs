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






public partial class WordEditorWindow : Window
{
    private readonly LoggingService _logger;
    private readonly WordEditorViewModel _viewModel;
    private readonly EditorState _editorState;

    private dynamic? _wordApp;
    private dynamic? _wordDocument;
    private bool _documentLoaded;
    private bool _closingConfirmed;

    
    
    
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
            
            if (!File.Exists(_editorState.LocalFilePath))
            {
                throw new FileNotFoundException($"File non trovato: {_editorState.LocalFilePath}");
            }

            
            
            LoadWordOnUiThread();

            
            await Task.Delay(500);

            
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

    
    
    
    
    private void LoadWordOnUiThread()
    {
        _logger.Log("Creazione istanza Word sul thread UI (STA)...");

        
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

        
        _wordApp.Visible = true;
        _wordApp.DisplayAlerts = 0; 

        
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
        
        if (_closingConfirmed)
        {
            CleanupWord();
            return;
        }

        
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
                    
                    if (SaveDocument(true))
                    {
                        DocumentReadyForConversion = _editorState.IsReadyForConversion;
                        CleanupWord();
                    }
                    else
                    {
                        
                        e.Cancel = true;
                    }
                    break;

                case MessageBoxResult.No:
                    
                    DocumentReadyForConversion = _editorState.IsReadyForConversion;
                    CleanupWord();
                    break;

                case MessageBoxResult.Cancel:
                    
                    e.Cancel = true;
                    break;
            }
        }
        else
        {
            
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




internal static class NativeMethods
{
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
}
