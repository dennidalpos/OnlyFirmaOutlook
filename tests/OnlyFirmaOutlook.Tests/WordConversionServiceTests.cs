using OnlyFirmaOutlook.Services;

namespace OnlyFirmaOutlook.Tests;

public class WordConversionServiceTests
{
    [Theory]
    [InlineData("Firma:Test", "Firma_Test")]
    [InlineData("   ", "Firma")]
    [InlineData("Nome__Firma", "Nome_Firma")]
    [InlineData("Firma@Test", "Firma@Test")]
    public void SanitizeFileName_ProducesSafeName(string input, string expected)
    {
        var sanitized = WordConversionService.SanitizeFileName(input);

        Assert.Equal(expected, sanitized);
    }

    [Fact]
    public void GenerateSignatureName_AppendsIdentifier()
    {
        var result = WordConversionService.GenerateSignatureName("Firma", "utente@example.com");

        Assert.Equal("Firma (utente@example.com)", result);
    }

    [Fact]
    public void Normalize_DoesNotThrowOnOfficePrefixedNodes()
    {
        var normalizer = new WordHtmlSignatureNormalizer();
        var html = """
            <html>
              <body>
                <o:p>&nbsp;</o:p>
                <p style="mso-margin-top-alt:auto; color:red;">Firma</p>
              </body>
            </html>
            """;

        var normalized = normalizer.Normalize(html);

        Assert.DoesNotContain("o:p", normalized, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("color:red", normalized, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("mso-margin-top-alt", normalized, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ProcessImages_HandlesPrefixedVmlNodesWithoutXPathErrors()
    {
        var tempDir = Path.Combine(Path.GetTempPath(), $"OnlyFirmaOutlookTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempDir);

        try
        {
            var imagePath = Path.Combine(tempDir, "logo.png");
            var sourceHtmlPath = Path.Combine(tempDir, "firma.htm");
            var assetsPath = Path.Combine(tempDir, "Firma_files");

            File.WriteAllBytes(imagePath, Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aBWQAAAAASUVORK5CYII="));
            File.WriteAllText(sourceHtmlPath, "<html></html>");

            var assetManager = new AssetManager();
            var html = $"""
                <html>
                  <body>
                    <v:imagedata src="{imagePath}" />
                    <v:shape o:href="{imagePath}"></v:shape>
                  </body>
                </html>
                """;

            var result = assetManager.ProcessImages(html, sourceHtmlPath, assetsPath, "Firma", useAbsolutePaths: false, embedImages: true);

            Assert.Contains("data:image/png;base64", result.Html, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            Directory.Delete(tempDir, recursive: true);
        }
    }
}
