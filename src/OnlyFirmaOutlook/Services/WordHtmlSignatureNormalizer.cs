using HtmlAgilityPack;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;

namespace OnlyFirmaOutlook.Services;

public class WordHtmlSignatureNormalizer
{
    public string Normalize(string html)
    {
        var doc = new HtmlDocument();
        doc.LoadHtml(html);

        var bodyNode = doc.DocumentNode.SelectSingleNode("//body");
        var workingHtml = bodyNode?.InnerHtml ?? doc.DocumentNode.InnerHtml;

        var workingDoc = new HtmlDocument();
        workingDoc.LoadHtml(workingHtml);

        RemoveUnsafeNodes(workingDoc);
        NormalizeNodes(workingDoc);

        var container = new HtmlDocument();
        container.LoadHtml("<div></div>");
        var root = container.DocumentNode.FirstChild;
        root.SetAttributeValue("style", "font-family:Calibri, Arial, sans-serif;font-size:12px;line-height:1.3;");
        root.InnerHtml = workingDoc.DocumentNode.InnerHtml;

        return root.OuterHtml;
    }

    private static void RemoveUnsafeNodes(HtmlDocument doc)
    {
        var nodesToRemove = doc.DocumentNode.SelectNodes("//script|//meta|//link|//xml|//style");
        if (nodesToRemove != null)
        {
            foreach (var node in nodesToRemove)
            {
                node.Remove();
            }
        }

        var commentNodes = doc.DocumentNode.SelectNodes("//comment()");
        if (commentNodes != null)
        {
            foreach (var comment in commentNodes)
            {
                comment.Remove();
            }
        }

        var namespaceNodes = doc.DocumentNode.Descendants()
            .Where(node => node.Name.StartsWith("o:", StringComparison.OrdinalIgnoreCase)
                || node.Name.StartsWith("v:", StringComparison.OrdinalIgnoreCase)
                || node.Name.StartsWith("w:", StringComparison.OrdinalIgnoreCase)
                || node.Name.Contains(':'));

        foreach (var node in namespaceNodes.ToList())
        {
            node.Remove();
        }
    }

    private static void NormalizeNodes(HtmlDocument doc)
    {
        foreach (var node in doc.DocumentNode.Descendants())
        {
            if (node.Attributes["style"] != null)
            {
                var cleaned = CleanStyle(node.Attributes["style"].Value);
                if (string.IsNullOrWhiteSpace(cleaned))
                {
                    node.Attributes.Remove("style");
                }
                else
                {
                    node.Attributes["style"].Value = cleaned;
                }
            }

            if (node.Name is "p" or "div" or "span")
            {
                AppendInlineStyle(node, "margin:0;");
            }
        }
    }

    private static string CleanStyle(string styleValue)
    {
        var parts = styleValue.Split(';', StringSplitOptions.RemoveEmptyEntries)
            .Select(part => part.Trim())
            .Where(part => !part.StartsWith("mso-", StringComparison.OrdinalIgnoreCase)
                && !part.Contains("tab-stops", StringComparison.OrdinalIgnoreCase)
                && !part.StartsWith("font-variant", StringComparison.OrdinalIgnoreCase));

        return string.Join("; ", parts);
    }

    private static void AppendInlineStyle(HtmlNode node, string styleFragment)
    {
        var existing = node.GetAttributeValue("style", string.Empty);
        if (existing.Contains("margin", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        var merged = string.IsNullOrWhiteSpace(existing)
            ? styleFragment
            : $"{existing.TrimEnd(';')}; {styleFragment}";

        node.SetAttributeValue("style", merged);
    }
}
