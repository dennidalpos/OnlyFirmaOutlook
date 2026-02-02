using System.IO;
using OnlyFirmaOutlook.Models;

namespace OnlyFirmaOutlook.Services;






public class WordEditorService
{
    private readonly LoggingService _logger;
    private readonly string _editorBaseTempFolder;

    public WordEditorService()
    {
        _logger = LoggingService.Instance;

        _editorBaseTempFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OnlyFirmaOutlook",
            "EditorTemp");

        try
        {
            Directory.CreateDirectory(_editorBaseTempFolder);
            _logger.Log($"Cartella base EditorTemp verificata: {_editorBaseTempFolder}");
        }
        catch (Exception ex)
        {
            _logger.LogError("Impossibile creare cartella EditorTemp base", ex);
            throw;
        }
    }

    
    
    
    
    
    
    public EditorState PrepareFileForEditing(string sourceFilePath, string proposedSignatureName)
    {
        if (!File.Exists(sourceFilePath))
        {
            throw new FileNotFoundException($"File sorgente non trovato: {sourceFilePath}");
        }

        _logger.Log($"Preparazione file per editing: {sourceFilePath}");

        var editorState = new EditorState
        {
            EditorSessionId = Guid.NewGuid(),
            ProposedSignatureName = proposedSignatureName
        };

        
        editorState.EditorTempFolder = Path.Combine(_editorBaseTempFolder, editorState.EditorSessionId.ToString());

        try
        {
            Directory.CreateDirectory(editorState.EditorTempFolder);
            _logger.Log($"Creata cartella editor temporanea: {editorState.EditorTempFolder}");
        }
        catch (Exception ex)
        {
            _logger.LogError("Impossibile creare cartella editor temporanea", ex);
            throw;
        }

        
        var fileName = Path.GetFileName(sourceFilePath);
        editorState.LocalFilePath = Path.Combine(editorState.EditorTempFolder, fileName);

        try
        {
            File.Copy(sourceFilePath, editorState.LocalFilePath, overwrite: true);

            
            File.SetAttributes(editorState.LocalFilePath, FileAttributes.Normal);

            _logger.Log($"File copiato in: {editorState.LocalFilePath}");
            editorState.LastModified = DateTime.Now;
        }
        catch (Exception ex)
        {
            _logger.LogError("Errore durante la copia del file per editing", ex);

            
            CleanupEditorTempFolder(editorState.EditorSessionId);
            throw;
        }

        return editorState;
    }

    
    
    
    
    public void CleanupEditorTempFolder(Guid editorSessionId)
    {
        var folderPath = Path.Combine(_editorBaseTempFolder, editorSessionId.ToString());

        if (!Directory.Exists(folderPath))
        {
            _logger.Log($"Cartella editor temporanea già eliminata o non esistente: {editorSessionId}");
            return;
        }

        TempCleanupHelper.CleanupDirectoryWithRetries(
            folderPath,
            _logger,
            $"cartella editor {editorSessionId}",
            retryDelayMs: 200);
    }

    
    
    
    
    public void CleanupOrphanedEditorFolders()
    {
        if (!Directory.Exists(_editorBaseTempFolder))
        {
            return;
        }

        _logger.Log("Verifica cartelle editor orfane...");

        try
        {
            foreach (var dir in Directory.GetDirectories(_editorBaseTempFolder))
            {
                var dirInfo = new DirectoryInfo(dir);

                
                if (dirInfo.LastWriteTime < DateTime.Now.AddDays(-1))
                {
                    try
                    {
                        Directory.Delete(dir, recursive: true);
                        _logger.Log($"Eliminata cartella editor orfana: {dirInfo.Name}");
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning($"Impossibile eliminare cartella editor orfana '{dir}': {ex.Message}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Errore durante pulizia cartelle editor orfane: {ex.Message}");
        }
    }

    public void CleanupAllEditorFolders()
    {
        if (!Directory.Exists(_editorBaseTempFolder))
        {
            return;
        }

        foreach (var dir in Directory.GetDirectories(_editorBaseTempFolder))
        {
            TempCleanupHelper.CleanupDirectoryWithRetries(
                dir,
                _logger,
                $"cartella editor {Path.GetFileName(dir)}",
                retryDelayMs: 200);
        }
    }

    
    
    
    public bool ValidateEditorState(EditorState editorState)
    {
        if (editorState == null)
        {
            return false;
        }

        return File.Exists(editorState.LocalFilePath);
    }

    
    
    
    public void UpdateLastModified(EditorState editorState)
    {
        editorState.LastModified = DateTime.Now;
    }
}
