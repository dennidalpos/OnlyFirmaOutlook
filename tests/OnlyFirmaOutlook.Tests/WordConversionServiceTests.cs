using OnlyFirmaOutlook.Services;

namespace OnlyFirmaOutlook.Tests;

public class WordConversionServiceTests
{
    [Theory]
    [InlineData("Firma:Test", "Firma_Test")]
    [InlineData("   ", "Firma")]
    [InlineData("Nome__Firma", "Nome_Firma")]
    public void SanitizeFileName_ProducesSafeName(string input, string expected)
    {
        var sanitized = WordConversionService.SanitizeFileName(input);

        Assert.Equal(expected, sanitized);
    }

    [Fact]
    public void GenerateSignatureName_AppendsIdentifier()
    {
        var result = WordConversionService.GenerateSignatureName("Firma", "utente@example.com");

        Assert.Equal("Firma (utente_example.com)", result);
    }
}
