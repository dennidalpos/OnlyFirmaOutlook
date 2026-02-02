using System.IO;

namespace OnlyFirmaOutlook.Services;

public class SignatureWorkflowService
{
    private readonly SignatureRepository _signatureRepository;
    private readonly WordConversionService _wordConversionService;
    private readonly LoggingService _logger;

    public SignatureWorkflowService(
        SignatureRepository signatureRepository,
        WordConversionService wordConversionService)
    {
        _signatureRepository = signatureRepository;
        _wordConversionService = wordConversionService;
        _logger = LoggingService.Instance;
    }

    public string BuildFinalSignatureName(string baseName, string? identifier)
    {
        return WordConversionService.GenerateSignatureName(baseName, identifier);
    }

    public bool ShouldCreateBackup(string destinationFolder)
    {
        if (string.IsNullOrWhiteSpace(destinationFolder))
        {
            return false;
        }

        var defaultFolder = SignatureRepository.GetDefaultOutlookSignaturesFolder();
        return string.Equals(
            Path.GetFullPath(destinationFolder).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
            Path.GetFullPath(defaultFolder).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
            StringComparison.OrdinalIgnoreCase);
    }

    public bool CreateBackupIfNeeded(string destinationFolder)
    {
        if (!ShouldCreateBackup(destinationFolder))
        {
            _logger.Log("Backup firme non necessario: cartella destinazione non predefinita.");
            return false;
        }

        return _signatureRepository.CreateBackupInSignaturesFolder();
    }

    public bool SignatureExists(string destinationFolder, string signatureName)
    {
        return _signatureRepository.SignatureExists(destinationFolder, signatureName);
    }

    public void DeleteExistingSignatureFiles(string destinationFolder, string signatureName)
    {
        _signatureRepository.DeleteExistingSignatureFiles(destinationFolder, signatureName);
    }

    public WordConversionService.ConversionResult ConvertDocument(
        string sourceDocPath,
        string destinationFolder,
        string signatureName,
        bool useFilteredHtml,
        bool fixOutlook2512 = true)
    {
        return _wordConversionService.ConvertDocument(sourceDocPath, destinationFolder, signatureName, useFilteredHtml, fixOutlook2512);
    }
}
