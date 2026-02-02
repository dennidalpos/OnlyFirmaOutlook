using System.ComponentModel;
using System.Runtime.CompilerServices;
using OnlyFirmaOutlook.Models;
using OnlyFirmaOutlook.Services;

namespace OnlyFirmaOutlook.ViewModels;





public class WordEditorViewModel : INotifyPropertyChanged
{
    private readonly LoggingService _logger;
    private readonly WordEditorService _editorService;
    private EditorState _editorState;

    private bool _isDocumentLoaded;
    private bool _hasUnsavedChanges;
    private bool _canSave;
    private string _statusText = "Pronto";

    public event PropertyChangedEventHandler? PropertyChanged;

    public EditorState EditorState
    {
        get => _editorState;
        set
        {
            _editorState = value;
            OnPropertyChanged();
        }
    }

    public bool IsDocumentLoaded
    {
        get => _isDocumentLoaded;
        set
        {
            if (_isDocumentLoaded != value)
            {
                _isDocumentLoaded = value;
                OnPropertyChanged();
                UpdateCommandStates();
            }
        }
    }

    public bool HasUnsavedChanges
    {
        get => _hasUnsavedChanges;
        set
        {
            if (_hasUnsavedChanges != value)
            {
                _hasUnsavedChanges = value;
                OnPropertyChanged();
                UpdateCommandStates();
                UpdateStatusText();
            }
        }
    }

    public bool CanSave
    {
        get => _canSave;
        set
        {
            if (_canSave != value)
            {
                _canSave = value;
                OnPropertyChanged();
            }
        }
    }

    public string StatusText
    {
        get => _statusText;
        set
        {
            if (_statusText != value)
            {
                _statusText = value;
                OnPropertyChanged();
            }
        }
    }

    public string WindowTitle => $"Editor Firma - {EditorState.ProposedSignatureName}";

    public WordEditorViewModel(EditorState editorState)
    {
        _logger = LoggingService.Instance;
        _editorService = new WordEditorService();
        _editorState = editorState;
    }

    public void MarkDocumentOpened()
    {
        EditorState.IsDocumentOpened = true;
        IsDocumentLoaded = true;
        _logger.Log("Documento aperto nell'editor");
        UpdateStatusText();
    }

    public void MarkDocumentSaved()
    {
        EditorState.IsDocumentSaved = true;
        EditorState.HasUnsavedChanges = false;
        HasUnsavedChanges = false;
        _editorService.UpdateLastModified(EditorState);
        _logger.Log("Documento salvato");
        UpdateStatusText();
    }

    public void MarkDocumentModified()
    {
        EditorState.HasUnsavedChanges = true;
        HasUnsavedChanges = true;
        UpdateStatusText();
    }

    private void UpdateCommandStates()
    {
        CanSave = IsDocumentLoaded;
    }

    private void UpdateStatusText()
    {
        StatusText = EditorState.GetStatusText();
    }

    protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
