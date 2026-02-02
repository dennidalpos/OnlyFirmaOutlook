using System.IO;
using OnlyFirmaOutlook.Models;

namespace OnlyFirmaOutlook.Services;




public class PresetService
{
    private readonly LoggingService _logger;
    private readonly string _mediaFolder;

    public PresetService()
    {
        _logger = LoggingService.Instance;
        _mediaFolder = Path.Combine(AppContext.BaseDirectory, "media");
    }

    
    
    
    public List<PresetFile> LoadPresets()
    {
        var presets = new List<PresetFile>();

        _logger.Log($"Ricerca preset in: {_mediaFolder}");

        if (!Directory.Exists(_mediaFolder))
        {
            _logger.LogWarning($"Cartella media non trovata: {_mediaFolder}");
            return presets;
        }

        try
        {
            var supportedExtensions = new[] { "*.doc", "*.docx", "*.rtf" };

            foreach (var pattern in supportedExtensions)
            {
                var files = Directory.GetFiles(_mediaFolder, pattern, SearchOption.TopDirectoryOnly);

                foreach (var file in files)
                {
                    
                    var fileName = Path.GetFileName(file);
                    if (fileName.StartsWith("~$", StringComparison.Ordinal))
                    {
                        continue;
                    }

                    var preset = new PresetFile
                    {
                        FullPath = file,
                        FileName = fileName,
                        DisplayName = Path.GetFileNameWithoutExtension(file)
                    };

                    presets.Add(preset);
                    _logger.Log($"Preset trovato: {preset.DisplayName}");
                }
            }

            _logger.Log($"Totale preset trovati: {presets.Count}");
        }
        catch (Exception ex)
        {
            _logger.LogError("Errore durante il caricamento dei preset", ex);
        }

        return presets.OrderBy(p => p.DisplayName).ToList();
    }

    
    
    
    public bool MediaFolderExists()
    {
        return Directory.Exists(_mediaFolder);
    }

    
    
    
    public string GetMediaFolderPath()
    {
        return _mediaFolder;
    }
}
