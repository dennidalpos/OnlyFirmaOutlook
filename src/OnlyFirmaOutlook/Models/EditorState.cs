namespace OnlyFirmaOutlook.Models;




public class EditorState
{
    
    
    
    public string LocalFilePath { get; set; } = string.Empty;

    
    
    
    public Guid EditorSessionId { get; set; }

    
    
    
    
    public string EditorTempFolder { get; set; } = string.Empty;

    
    
    
    public bool IsDocumentOpened { get; set; }

    
    
    
    public bool IsDocumentSaved { get; set; }

    
    
    
    public bool HasUnsavedChanges { get; set; }

    
    
    
    public string ProposedSignatureName { get; set; } = string.Empty;

    
    
    
    public DateTime LastModified { get; set; }

    
    
    
    
    public bool IsReadyForConversion => IsDocumentOpened && IsDocumentSaved;

    public EditorState()
    {
        EditorSessionId = Guid.NewGuid();
        LastModified = DateTime.Now;
    }

    
    
    
    public string GetStatusText()
    {
        if (!IsDocumentOpened)
            return "Da modificare";

        if (!IsDocumentSaved)
            return "Aperto ma non salvato";

        if (HasUnsavedChanges)
            return "Modificato (non salvato)";

        return "Modificata e pronta";
    }
}
