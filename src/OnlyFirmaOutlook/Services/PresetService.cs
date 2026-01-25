using System.IO;
using OnlyFirmaOutlook.Models;

namespace OnlyFirmaOutlook.Services;

/// <summary>
/// Servizio per la gestione dei file preset nella cartella media.
/// </summary>
public class PresetService
{
    private readonly LoggingService _logger;
    private readonly string _mediaFolder;

    public PresetService()
    {
        _logger = LoggingService.Instance;
        _mediaFolder = Path.Combine(AppContext.BaseDirectory, "media");
    }

    /// <summary>
    /// Carica i file preset dalla cartella media.
    /// </summary>
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
            // Cerca file Word (.doc e .docx)
            var wordExtensions = new[] { "*.doc", "*.docx" };

            foreach (var pattern in wordExtensions)
            {
                var files = Directory.GetFiles(_mediaFolder, pattern, SearchOption.TopDirectoryOnly);

                foreach (var file in files)
                {
                    // Salta file temporanei di Word (iniziano con ~$)
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

    /// <summary>
    /// Verifica se la cartella media esiste.
    /// </summary>
    public bool MediaFolderExists()
    {
        return Directory.Exists(_mediaFolder);
    }

    /// <summary>
    /// Ottiene il percorso della cartella media.
    /// </summary>
    public string GetMediaFolderPath()
    {
        return _mediaFolder;
    }
}
