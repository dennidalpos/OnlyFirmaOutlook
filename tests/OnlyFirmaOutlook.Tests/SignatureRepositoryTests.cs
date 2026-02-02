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
}
