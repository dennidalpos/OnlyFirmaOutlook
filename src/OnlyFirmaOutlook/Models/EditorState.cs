namespace OnlyFirmaOutlook.Models;

/// <summary>
/// Rappresenta lo stato di modifica di un documento nell'editor Word.
/// </summary>
public class EditorState
{
    /// <summary>
    /// Percorso del file Word locale (copia temporanea in EditorTemp).
    /// </summary>
    public string LocalFilePath { get; set; } = string.Empty;

    /// <summary>
    /// GUID della cartella temporanea dedicata per questa sessione di editing.
    /// </summary>
    public Guid EditorSessionId { get; set; }

    /// <summary>
    /// Percorso completo della cartella temporanea.
    /// %LOCALAPPDATA%\OnlyFirmaOutlook\EditorTemp\{guid}
    /// </summary>
    public string EditorTempFolder { get; set; } = string.Empty;

    /// <summary>
    /// Indica se il documento è stato aperto nell'editor almeno una volta.
    /// </summary>
    public bool IsDocumentOpened { get; set; }

    /// <summary>
    /// Indica se il documento è stato salvato almeno una volta.
    /// </summary>
    public bool IsDocumentSaved { get; set; }

    /// <summary>
    /// Indica se il documento ha modifiche non salvate.
    /// </summary>
    public bool HasUnsavedChanges { get; set; }

    /// <summary>
    /// Nome base per la firma (proposto dall'utente o dal preset).
    /// </summary>
    public string ProposedSignatureName { get; set; } = string.Empty;

    /// <summary>
    /// Timestamp ultima modifica.
    /// </summary>
    public DateTime LastModified { get; set; }

    /// <summary>
    /// Verifica se il documento è pronto per la conversione.
    /// Requisito: deve essere stato aperto E salvato almeno una volta.
    /// </summary>
    public bool IsReadyForConversion => IsDocumentOpened && IsDocumentSaved;

    public EditorState()
    {
        EditorSessionId = Guid.NewGuid();
        LastModified = DateTime.Now;
    }

    /// <summary>
    /// Restituisce lo stato come stringa descrittiva.
    /// </summary>
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
