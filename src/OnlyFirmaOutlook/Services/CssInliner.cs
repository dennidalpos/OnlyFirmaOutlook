using HtmlAgilityPack;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;
using PreMailer.Net;

namespace OnlyFirmaOutlook.Services;

public class CssInliner
{
    public string InlineCss(string html)
    {
        try
        {
            var wrapped = WrapHtml(html);
            var result = PreMailer.Net.PreMailer.MoveCssInline(wrapped, removeStyleElements: true);
            return ExtractBodyHtml(result.Html);
        }
        catch
        {
            return FallbackInline(html);
        }
    }

    private static string WrapHtml(string html)
    {
        if (html.Contains("<html", StringComparison.OrdinalIgnoreCase))
        {
            return html;
        }

        return $"<html><head></head><body>{html}</body></html>";
    }

    private static string ExtractBodyHtml(string html)
    {
        var doc = new HtmlDocument();
        doc.LoadHtml(html);
        var bodyNode = doc.DocumentNode.SelectSingleNode("//body");
        return bodyNode?.InnerHtml ?? doc.DocumentNode.InnerHtml;
    }

    private static string FallbackInline(string html)
    {
        var doc = new HtmlDocument();
        doc.LoadHtml(html);
        var styleNodes = doc.DocumentNode.SelectNodes("//style");
        if (styleNodes != null)
        {
            foreach (var node in styleNodes)
            {
                node.Remove();
            }
        }

        foreach (var node in doc.DocumentNode.Descendants())
        {
            node.Attributes.Remove("class");
        }

        return doc.DocumentNode.InnerHtml;
    }
}
