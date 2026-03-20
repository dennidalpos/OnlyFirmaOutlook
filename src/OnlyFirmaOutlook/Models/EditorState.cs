/*
 * OnlyFirmaOutlook
 * Copyright (c) 2026 Danny Perondi. All rights reserved.
 * Author: Danny Perondi
 * Proprietary and confidential.
 * Unauthorized copying, modification, distribution, sublicensing, disclosure,
 * or commercial use is prohibited without prior written permission.
 */

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

    
    
    
    
    public bool IsReadyForConversion => !IsDocumentOpened && IsDocumentSaved && !HasUnsavedChanges;

    public EditorState()
    {
        EditorSessionId = Guid.NewGuid();
        LastModified = DateTime.Now;
    }

    
    
    
    public string GetStatusText()
    {
        if (IsDocumentOpened)
        {
            if (HasUnsavedChanges)
                return IsDocumentSaved ? "Modificato (non salvato)" : "Aperto ma non salvato";

            if (IsDocumentSaved)
                return "Salvato: chiudi Word";
        }

        if (IsDocumentSaved && !HasUnsavedChanges)
            return "Modificata e pronta";

        return "Da modificare";
    }
}
