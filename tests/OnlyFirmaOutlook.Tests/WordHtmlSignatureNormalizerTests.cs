using OnlyFirmaOutlook.Services;

namespace OnlyFirmaOutlook.Tests;

public class WordHtmlSignatureNormalizerTests
{
    [Fact]
    public void Normalize_PreservesHiddenTableStyles()
    {
        var html = "<table style=\"mso-hide:all\" border=\"1\"><tr><td style=\"border:1px solid #000\">Hidden</td></tr></table>";
        var normalizer = new WordHtmlSignatureNormalizer();

        var result = normalizer.Normalize(html);

        Assert.Contains("mso-hide:all", result);
        Assert.Contains("display:none", result);
        Assert.DoesNotContain("border=\"1\"", result);
        Assert.DoesNotContain("border:1px", result);
    }
}
