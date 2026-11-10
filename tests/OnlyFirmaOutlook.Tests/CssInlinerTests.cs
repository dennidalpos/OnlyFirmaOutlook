using OnlyFirmaOutlook.Services;

namespace OnlyFirmaOutlook.Tests;

public class CssInlinerTests
{
    private readonly CssInliner _inliner;

    public CssInlinerTests()
    {
        _inliner = new CssInliner();
    }

    [Fact]
    public void InlineCss_InlinesStyleTagsToAttributes()
    {
        var input = "<style>.test { color: red; }</style><div class=\"test\">Hello</div>";
        var expectedContent = "color: red";
        
        var result = _inliner.InlineCss(input);
        
        Assert.Contains(expectedContent, result);
        Assert.DoesNotContain("<style>", result);
    }

    [Fact]
    public void InlineCss_ExtractsBodyContent_WhenWrapped()
    {
        var input = "<html><head><style>.test { color: blue; }</style></head><body><div class=\"test\">Hello</div></body></html>";
        
        var result = _inliner.InlineCss(input);
        
        Assert.Contains("color: blue", result);
        Assert.DoesNotContain("<html>", result);
        Assert.DoesNotContain("<body>", result);
    }

    [Fact]
    public void InlineCss_PreservesHtmlWithoutStyles()
    {
        var input = "<p>No styles here</p>";
        var result = _inliner.InlineCss(input);
        Assert.Equal(input, result);
    }

    [Fact]
    public void InlineCss_HandlesEmptyInput()
    {
        var result = _inliner.InlineCss("");
        Assert.True(string.IsNullOrWhiteSpace(result));
    }

    [Fact]
    public void InlineCss_HandlesComplexCssWithMultipleSelectors()
    {
        var input = """
            <style>
                .red-text { color: red; }
                .bold-text { font-weight: bold; }
            </style>
            <p class="red-text bold-text">Hello</p>
            """;
            
        var result = _inliner.InlineCss(input);
        
        Assert.Contains("color: red", result);
        Assert.Contains("font-weight: bold", result);
        Assert.DoesNotContain("<style>", result);
    }
}
