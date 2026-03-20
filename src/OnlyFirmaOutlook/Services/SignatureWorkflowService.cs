/*
 * OnlyFirmaOutlook
 * Copyright (c) 2026 Danny Perondi. All rights reserved.
 * Author: Danny Perondi
 * Proprietary and confidential.
 * Unauthorized copying, modification, distribution, sublicensing, disclosure,
 * or commercial use is prohibited without prior written permission.
 */

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

    public bool ShouldCreateBackupBeforeOverwrite(string destinationFolder, bool signatureExists)
    {
        return signatureExists && ShouldCreateBackup(destinationFolder);
    }

    public bool CreateBackupIfNeeded(string destinationFolder, bool signatureExists)
    {
        if (!ShouldCreateBackupBeforeOverwrite(destinationFolder, signatureExists))
        {
            _logger.Log("Backup firme non necessario: nessuna sovrascrittura nella cartella predefinita.");
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
        bool useFilteredHtml)
    {
        return _wordConversionService.ConvertDocument(sourceDocPath, destinationFolder, signatureName, useFilteredHtml);
    }
}
