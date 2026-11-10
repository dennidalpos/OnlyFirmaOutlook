using OnlyFirmaOutlook.Services;

namespace OnlyFirmaOutlook.Tests;

public class SignatureWorkflowServiceTests
{
    private readonly SignatureWorkflowService _workflowService = new(
        new SignatureRepository(),
        new WordConversionService());

    [Fact]
    public void BuildFinalSignatureName_AppendsIdentifierWhenProvided()
    {
        var result = _workflowService.BuildFinalSignatureName("Firma", "utente@example.com");

        Assert.Equal("Firma (utente@example.com)", result);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void ShouldCreateBackup_ReturnsFalseForEmptyDestination(string? destinationFolder)
    {
        var shouldCreateBackup = _workflowService.ShouldCreateBackup(destinationFolder!);

        Assert.False(shouldCreateBackup);
    }

    [Fact]
    public void ShouldCreateBackup_ReturnsTrueForDefaultOutlookFolderWithTrailingSeparator()
    {
        var defaultFolder = SignatureRepository.GetDefaultOutlookSignaturesFolder();
        var destinationFolder = defaultFolder + Path.DirectorySeparatorChar;

        var shouldCreateBackup = _workflowService.ShouldCreateBackup(destinationFolder);

        Assert.True(shouldCreateBackup);
    }

    [Fact]
    public void ShouldCreateBackup_ReturnsFalseForAlternativeFolder()
    {
        var destinationFolder = SignatureRepository.GetAlternativeOutputFolder();

        var shouldCreateBackup = _workflowService.ShouldCreateBackup(destinationFolder);

        Assert.False(shouldCreateBackup);
    }

    [Fact]
    public void ShouldCreateBackupBeforeOverwrite_ReturnsFalseWhenSignatureDoesNotExist()
    {
        var destinationFolder = SignatureRepository.GetDefaultOutlookSignaturesFolder();

        var shouldCreateBackup = _workflowService.ShouldCreateBackupBeforeOverwrite(destinationFolder, signatureExists: false);

        Assert.False(shouldCreateBackup);
    }

    [Fact]
    public void ShouldCreateBackupBeforeOverwrite_ReturnsTrueWhenOverwritingInDefaultFolder()
    {
        var destinationFolder = SignatureRepository.GetDefaultOutlookSignaturesFolder();

        var shouldCreateBackup = _workflowService.ShouldCreateBackupBeforeOverwrite(destinationFolder, signatureExists: true);

        Assert.True(shouldCreateBackup);
    }

    [Fact]
    public void CreateBackupIfNeeded_ReturnsFalseWhenDestinationIsNotDefaultFolder()
    {
        var destinationFolder = SignatureRepository.GetAlternativeOutputFolder();

        var created = _workflowService.CreateBackupIfNeeded(destinationFolder, signatureExists: true);

        Assert.False(created);
    }

    [Fact]
    public void CreateBackupIfNeeded_ReturnsFalseWhenSignatureDoesNotExist()
    {
        var destinationFolder = SignatureRepository.GetDefaultOutlookSignaturesFolder();

        var created = _workflowService.CreateBackupIfNeeded(destinationFolder, signatureExists: false);

        Assert.False(created);
    }
}
