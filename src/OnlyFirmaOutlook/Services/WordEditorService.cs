using System.IO;
using OnlyFirmaOutlook.Helpers;
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

        _logger.Log($"Avvio cleanup cartella editor: {folderPath}");

        var maxRetries = 3;
        var retryDelayMs = 200;

        for (var attempt = 1; attempt <= maxRetries; attempt++)
        {
            try
            {
                
                foreach (var file in Directory.GetFiles(folderPath, "*", SearchOption.AllDirectories))
                {
                    try
                    {
                        File.SetAttributes(file, FileAttributes.Normal);
                        File.Delete(file);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning($"Impossibile eliminare file '{file}': {ex.Message}");
                    }
                }

                
                foreach (var dir in Directory.GetDirectories(folderPath, "*", SearchOption.AllDirectories)
                                             .OrderByDescending(d => d.Length))
                {
                    try
                    {
                        Directory.Delete(dir, recursive: false);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning($"Impossibile eliminare cartella '{dir}': {ex.Message}");
                    }
                }

                
                Directory.Delete(folderPath, recursive: true);
                _logger.Log($"Cleanup cartella editor completato: {editorSessionId}");
                return;
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Tentativo {attempt}/{maxRetries} cleanup editor fallito: {ex.Message}");

                if (attempt < maxRetries)
                {
                    Thread.Sleep(retryDelayMs * attempt);
                }
            }
        }

        _logger.LogWarning($"Cleanup cartella editor non completato dopo {maxRetries} tentativi. " +
                          $"La cartella '{folderPath}' potrebbe richiedere pulizia manuale.");
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
