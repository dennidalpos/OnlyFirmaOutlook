/*
 * OnlyFirmaOutlook
 * Copyright (c) 2026 Danny Perondi. All rights reserved.
 * Author: Danny Perondi
 * Proprietary and confidential.
 * Unauthorized copying, modification, distribution, sublicensing, disclosure,
 * or commercial use is prohibited without prior written permission.
 */

using OnlyFirmaOutlook.Models;

namespace OnlyFirmaOutlook.Services;

internal static class EditorStateTransitions
{
    internal static void MarkDocumentOpened(EditorState editorState)
    {
        editorState.IsDocumentOpened = true;
        editorState.HasUnsavedChanges = true;
    }

    internal static void MarkDocumentClosed(EditorState editorState)
    {
        editorState.IsDocumentOpened = false;
    }

    internal static bool TryMarkDocumentSaved(
        EditorState editorState,
        DateTime observedModifiedTime,
        ref DateTime lastKnownModifiedTime)
    {
        if (observedModifiedTime <= lastKnownModifiedTime)
        {
            return false;
        }

        lastKnownModifiedTime = observedModifiedTime;
        editorState.IsDocumentSaved = true;
        editorState.HasUnsavedChanges = false;
        editorState.LastModified = observedModifiedTime;
        return true;
    }
}
