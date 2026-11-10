using OnlyFirmaOutlook.Services;

namespace OnlyFirmaOutlook.Tests;

public class SignatureRepositoryTests
{
    [Fact]
    public void GetSignatures_IncludesRtfAndTxtOnlyEntries()
    {
        var tempDir = Path.Combine(Path.GetTempPath(), $"OnlyFirmaOutlookTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempDir);

        try
        {
            File.WriteAllText(Path.Combine(tempDir, "FirmaSoloRtf.rtf"), "rtf");
            File.WriteAllText(Path.Combine(tempDir, "FirmaSoloTxt.txt"), "txt");
            File.WriteAllText(Path.Combine(tempDir, "FirmaCompleta.htm"), "html");
            File.WriteAllText(Path.Combine(tempDir, "FirmaCompleta.rtf"), "rtf");
            Directory.CreateDirectory(Path.Combine(tempDir, "FirmaCompleta_files"));

            var repository = new SignatureRepository();

            var signatures = repository.GetSignatures(tempDir);

            Assert.Contains(signatures, signature => signature.Name == "FirmaSoloRtf" && signature.HasRtf);
            Assert.Contains(signatures, signature => signature.Name == "FirmaSoloTxt" && signature.HasTxt);
            Assert.Contains(signatures, signature => signature.Name == "FirmaCompleta" && signature.HasHtm && signature.HasRtf);
        }
        finally
        {
            Directory.Delete(tempDir, recursive: true);
        }
    }

    [Fact]
    public void GetSignatures_IncludesAssetOnlyEntries()
    {
        var tempDir = Path.Combine(Path.GetTempPath(), $"OnlyFirmaOutlookTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempDir);

        try
        {
            Directory.CreateDirectory(Path.Combine(tempDir, "FirmaAssets_files"));

            var repository = new SignatureRepository();

            var signatures = repository.GetSignatures(tempDir);

            Assert.Contains(signatures, signature => signature.Name == "FirmaAssets" && signature.HasFilesFolder);
        }
        finally
        {
            Directory.Delete(tempDir, recursive: true);
        }
    }

    [Theory]
    [InlineData("FirmaRtf", ".rtf")]
    [InlineData("FirmaTxt", ".txt")]
    public void SignatureExists_RecognizesDocumentArtifacts(string signatureName, string extension)
    {
        var tempDir = Path.Combine(Path.GetTempPath(), $"OnlyFirmaOutlookTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempDir);

        try
        {
            File.WriteAllText(Path.Combine(tempDir, signatureName + extension), "artifact");

            var repository = new SignatureRepository();

            Assert.True(repository.SignatureExists(tempDir, signatureName));
        }
        finally
        {
            Directory.Delete(tempDir, recursive: true);
        }
    }

    [Theory]
    [InlineData("FirmaAssets_files")]
    [InlineData("FirmaAssets_file")]
    public void SignatureExists_RecognizesAssetFolders(string folderName)
    {
        var tempDir = Path.Combine(Path.GetTempPath(), $"OnlyFirmaOutlookTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempDir);

        try
        {
            Directory.CreateDirectory(Path.Combine(tempDir, folderName));

            var repository = new SignatureRepository();

            Assert.True(repository.SignatureExists(tempDir, "FirmaAssets"));
        }
        finally
        {
            Directory.Delete(tempDir, recursive: true);
        }
    }

    [Fact]
    public void RestoreBackup_RemovesArtifactsNotPresentInBackupBeforeExtract()
    {
        var tempDir = Path.Combine(Path.GetTempPath(), $"OnlyFirmaOutlookTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempDir);

        try
        {
            File.WriteAllText(Path.Combine(tempDir, "StaleSignature.htm"), "stale");
            Directory.CreateDirectory(Path.Combine(tempDir, "StaleSignature_files"));
            File.WriteAllText(Path.Combine(tempDir, "StaleSignature_files", "old.png"), "old");

            var backupPath = Path.Combine(tempDir, "backup_firme_onlyfirmaoutlook_test.zip");
            using (var archive = System.IO.Compression.ZipFile.Open(backupPath, System.IO.Compression.ZipArchiveMode.Create))
            {
                var htmlEntry = archive.CreateEntry("RestoredSignature.htm");
                using var htmlWriter = new StreamWriter(htmlEntry.Open());
                htmlWriter.Write("restored");
            }

            var repository = new SignatureRepository();
            var backup = new OnlyFirmaOutlook.Models.BackupInfo
            {
                FileName = Path.GetFileName(backupPath),
                FullPath = backupPath
            };

            var restored = repository.RestoreBackup(backup, tempDir);

            Assert.True(restored);
            Assert.False(File.Exists(Path.Combine(tempDir, "StaleSignature.htm")));
            Assert.False(Directory.Exists(Path.Combine(tempDir, "StaleSignature_files")));
            Assert.True(File.Exists(Path.Combine(tempDir, "RestoredSignature.htm")));
            Assert.True(File.Exists(backupPath));
        }
        finally
        {
            Directory.Delete(tempDir, recursive: true);
        }
    }
}
