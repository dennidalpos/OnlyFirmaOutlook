using OnlyFirmaOutlook.Models;
using OnlyFirmaOutlook.Services;

namespace OnlyFirmaOutlook.Tests;

public class EditorStateTransitionsTests
{
    [Fact]
    public void MarkDocumentOpened_SetsOpenedFlag()
    {
        var editorState = new EditorState
        {
            IsDocumentOpened = false,
            IsDocumentSaved = false,
            HasUnsavedChanges = false
        };

        EditorStateTransitions.MarkDocumentOpened(editorState);

        Assert.True(editorState.IsDocumentOpened);
        Assert.True(editorState.HasUnsavedChanges);
        Assert.False(editorState.IsReadyForConversion);
    }

    [Fact]
    public void MarkDocumentClosed_ClearsOpenedFlag()
    {
        var editorState = new EditorState
        {
            IsDocumentOpened = true,
            IsDocumentSaved = true,
            HasUnsavedChanges = false
        };

        EditorStateTransitions.MarkDocumentClosed(editorState);

        Assert.False(editorState.IsDocumentOpened);
        Assert.True(editorState.IsReadyForConversion);
    }

    [Fact]
    public void TryMarkDocumentSaved_UpdatesStateWhenTimestampIsNewer()
    {
        var initialTimestamp = new DateTime(2026, 03, 19, 10, 00, 00, DateTimeKind.Local);
        var observedTimestamp = initialTimestamp.AddMinutes(5);
        var lastKnownModifiedTime = initialTimestamp;
        var editorState = new EditorState
        {
            IsDocumentOpened = true,
            IsDocumentSaved = false,
            HasUnsavedChanges = true,
            LastModified = initialTimestamp
        };

        var changed = EditorStateTransitions.TryMarkDocumentSaved(
            editorState,
            observedTimestamp,
            ref lastKnownModifiedTime);

        Assert.True(changed);
        Assert.True(editorState.IsDocumentSaved);
        Assert.False(editorState.HasUnsavedChanges);
        Assert.False(editorState.IsReadyForConversion);
        Assert.Equal(observedTimestamp, editorState.LastModified);
        Assert.Equal(observedTimestamp, lastKnownModifiedTime);
    }

    [Fact]
    public void TryMarkDocumentSaved_IgnoresUnchangedTimestamp()
    {
        var initialTimestamp = new DateTime(2026, 03, 19, 10, 00, 00, DateTimeKind.Local);
        var lastKnownModifiedTime = initialTimestamp;
        var editorState = new EditorState
        {
            IsDocumentOpened = true,
            IsDocumentSaved = false,
            LastModified = initialTimestamp
        };

        var changed = EditorStateTransitions.TryMarkDocumentSaved(
            editorState,
            initialTimestamp,
            ref lastKnownModifiedTime);

        Assert.False(changed);
        Assert.False(editorState.IsDocumentSaved);
        Assert.Equal(initialTimestamp, editorState.LastModified);
        Assert.Equal(initialTimestamp, lastKnownModifiedTime);
    }

    [Theory]
    [InlineData(false, false, false, "Da modificare")]
    [InlineData(true, false, true, "Aperto ma non salvato")]
    [InlineData(true, true, true, "Modificato (non salvato)")]
    [InlineData(true, true, false, "Salvato: chiudi Word")]
    [InlineData(false, true, false, "Modificata e pronta")]
    public void GetStatusText_ReturnsExpectedValue(
        bool isDocumentOpened,
        bool isDocumentSaved,
        bool hasUnsavedChanges,
        string expectedStatus)
    {
        var editorState = new EditorState
        {
            IsDocumentOpened = isDocumentOpened,
            IsDocumentSaved = isDocumentSaved,
            HasUnsavedChanges = hasUnsavedChanges
        };

        var status = editorState.GetStatusText();

        Assert.Equal(expectedStatus, status);
    }
}
