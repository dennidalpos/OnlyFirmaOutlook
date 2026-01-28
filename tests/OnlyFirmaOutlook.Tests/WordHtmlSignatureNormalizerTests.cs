using OnlyFirmaOutlook.Services;

namespace OnlyFirmaOutlook.Tests;

public class WordHtmlSignatureNormalizerTests
{
    [Fact]
    public void Normalize_PreservesHiddenTableStyles()
    {
        var html = "<table style=\"mso-hide:all\" border=\"1\"><tr><td>Hidden</td></tr></table>";
        var normalizer = new WordHtmlSignatureNormalizer();

        var result = normalizer.Normalize(html);

        Assert.Contains("mso-hide:all", result);
        Assert.Contains("display:none", result);
        Assert.DoesNotContain("border=\"1\"", result);
    }
}
