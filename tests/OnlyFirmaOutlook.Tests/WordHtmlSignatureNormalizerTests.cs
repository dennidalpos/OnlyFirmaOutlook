using OnlyFirmaOutlook.Services;

namespace OnlyFirmaOutlook.Tests;

public class WordHtmlSignatureNormalizerTests
{
    private readonly WordHtmlSignatureNormalizer _normalizer;

    public WordHtmlSignatureNormalizerTests()
    {
        _normalizer = new WordHtmlSignatureNormalizer();
    }

    [Theory]
    [InlineData("<script>alert('xss');</script><p>test</p>", "<p>test</p>")]
    [InlineData("<meta charset=\"utf-8\"><p>test</p>", "<p>test</p>")]
    [InlineData("<xml><foo>bar</foo></xml><p>test</p>", "<p>test</p>")]
    [InlineData("<p>test<o:p>&nbsp;</o:p></p>", "<p>test</p>")]
    [InlineData("<!-- comment --><p>test</p>", "<p>test</p>")]
    [InlineData("<w:WordDocument><w:View>Print</w:View></w:WordDocument><p>test</p>", "<p>test</p>")]
    public void Normalize_RemovesNonRenderingElements(string input, string expected)
    {
        var result = _normalizer.Normalize(input);
        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("<p style=\"mso-bidi-font-family: 'Times New Roman'; font-family: Arial;\">test</p>", "<p style=\"font-family: Arial\">test</p>")]
    [InlineData("<p style=\"tab-stops: 36.0pt; color: red;\">test</p>", "<p style=\"color: red\">test</p>")]
    [InlineData("<p style=\"mso-line-height-rule: exactly; font-size: 10pt;\">test</p>", "<p style=\"mso-line-height-rule: exactly; font-size: 10pt\">test</p>")]
    public void Normalize_CleansCssStyles_WhilePreservingOthers(string input, string expected)
    {
        var result = _normalizer.Normalize(input);
        Assert.Equal(expected, result);
    }
    
    [Fact]
    public void Normalize_ExtractsBodyContent_WhenPresent()
    {
        var input = "<html><head><style>p { color: red; }</style></head><body><p>test</p></body></html>";
        var result = _normalizer.Normalize(input);
        Assert.Equal("<p>test</p>", result);
    }
    
    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void Normalize_HandlesEmptyInput(string input)
    {
        var result = _normalizer.Normalize(input);
        Assert.True(string.IsNullOrWhiteSpace(result));
    }

    [Fact]
    public void Normalize_HandlesComplexHtml()
    {
        var input = """
            <html>
            <body>
                <!-- HTML Comment -->
                <script>doBadThings();</script>
                <p style="mso-margin-top-alt: auto; font-family: 'Arial'; mso-line-height-rule: exactly;">
                    Hello World<w:Sdt></w:Sdt><o:p></o:p>
                </p>
                <xml><foo></foo></xml>
            </body>
            </html>
            """;
            
        var result = _normalizer.Normalize(input);
        
        Assert.Contains("font-family: 'Arial'", result);
        Assert.Contains("mso-line-height-rule: exactly", result);
        Assert.Contains("Hello World", result);
        Assert.DoesNotContain("<!--", result);
        Assert.DoesNotContain("<script>", result);
        Assert.DoesNotContain("mso-margin-top-alt", result);
        Assert.DoesNotContain("<w:", result);
        Assert.DoesNotContain("<o:p>", result);
        Assert.DoesNotContain("<xml>", result);
    }
}
