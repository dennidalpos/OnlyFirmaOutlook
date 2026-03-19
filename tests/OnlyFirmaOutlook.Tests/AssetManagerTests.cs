using OnlyFirmaOutlook.Services;

namespace OnlyFirmaOutlook.Tests;

public class AssetManagerTests
{
    [Theory]
    [InlineData("image/webp", ".webp")]
    [InlineData("image/x-icon", ".ico")]
    public void ProcessImages_SavesEmbeddedDataUrisWithExpectedExtension(string mimeType, string expectedExtension)
    {
        var tempDir = Path.Combine(Path.GetTempPath(), $"OnlyFirmaOutlookTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempDir);

        try
        {
            var assetsPath = Path.Combine(tempDir, "Firma_files");
            var sourceHtmlPath = Path.Combine(tempDir, "firma.htm");
            File.WriteAllText(sourceHtmlPath, "<html></html>");

            var assetManager = new AssetManager();
            var html = $"""
                <html>
                  <body>
                    <img src="data:{mimeType};base64,AA==" />
                  </body>
                </html>
                """;

            var result = assetManager.ProcessImages(
                html,
                sourceHtmlPath,
                assetsPath,
                "Firma",
                useAbsolutePaths: false,
                embedImages: false);

            var savedFile = Directory.GetFiles(assetsPath, "*" + expectedExtension, SearchOption.TopDirectoryOnly);

            Assert.Single(savedFile);
            Assert.Contains($"Firma_files/", result.Html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain(".img", result.Html, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            Directory.Delete(tempDir, recursive: true);
        }
    }
}
