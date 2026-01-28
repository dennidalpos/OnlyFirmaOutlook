using System;
using System.Collections.Generic;
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
        NormalizeHiddenTables(workingDoc);

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

        var commentNodes = doc.DocumentNode.Descendants()
            .Where(node => node.NodeType == HtmlNodeType.Comment)
            .ToList();

        foreach (var comment in commentNodes)
        {
            comment.Remove();
        }

        var namespaceNodes = doc.DocumentNode.Descendants()
            .Where(node => node.Name.StartsWith("w:", StringComparison.OrdinalIgnoreCase)
                           || node.Name.StartsWith("o:", StringComparison.OrdinalIgnoreCase))
            .ToList();

        foreach (var node in namespaceNodes)
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
                AppendInlineStyleIfMissing(node, "margin:0;", "margin");
            }
        }
    }

    private static void NormalizeHiddenTables(HtmlDocument doc)
    {
        var hiddenNodes = doc.DocumentNode.Descendants()
            .Where(HasHiddenStyle)
            .ToList();

        foreach (var node in hiddenNodes)
        {
            var table = node.Name.Equals("table", StringComparison.OrdinalIgnoreCase)
                ? node
                : node.Ancestors("table").FirstOrDefault();

            if (table == null)
            {
                continue;
            }

            EnsureDisplayNone(table);
            RemoveTableBorders(table);
        }
    }

    private static string CleanStyle(string styleValue)
    {
        var parts = styleValue.Split(';', StringSplitOptions.RemoveEmptyEntries)
            .Select(part => part.Trim());

        var cleanedParts = new List<string>();
        var hasMsoHide = false;
        var hasDisplayNone = false;

        foreach (var part in parts)
        {
            if (part.StartsWith("mso-hide", StringComparison.OrdinalIgnoreCase))
            {
                hasMsoHide = true;
                cleanedParts.Add(part);
                continue;
            }

            if (part.StartsWith("display", StringComparison.OrdinalIgnoreCase)
                && part.Contains("none", StringComparison.OrdinalIgnoreCase))
            {
                hasDisplayNone = true;
            }

            if (part.StartsWith("mso-", StringComparison.OrdinalIgnoreCase)
                || part.Contains("tab-stops", StringComparison.OrdinalIgnoreCase)
                || part.StartsWith("font-variant", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            cleanedParts.Add(part);
        }

        if (hasMsoHide && !hasDisplayNone)
        {
            cleanedParts.Add("display:none");
        }

        return string.Join("; ", cleanedParts);
    }

    private static bool HasHiddenStyle(HtmlNode node)
    {
        var style = node.GetAttributeValue("style", string.Empty);
        if (!string.IsNullOrWhiteSpace(style))
        {
            if (style.Contains("mso-hide", StringComparison.OrdinalIgnoreCase)
                || style.Contains("display:none", StringComparison.OrdinalIgnoreCase)
                || style.Contains("visibility:hidden", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        if (node.Attributes.Contains("hidden"))
        {
            return true;
        }

        return false;
    }

    private static void EnsureDisplayNone(HtmlNode node)
    {
        var style = node.GetAttributeValue("style", string.Empty);
        if (style.Contains("display:none", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        var merged = string.IsNullOrWhiteSpace(style)
            ? "display:none"
            : $"{style.TrimEnd(';')}; display:none";

        node.SetAttributeValue("style", merged);
    }

    private static void RemoveTableBorders(HtmlNode table)
    {
        table.Attributes.Remove("border");
        table.Attributes.Remove("cellspacing");
        table.Attributes.Remove("cellpadding");

        var style = table.GetAttributeValue("style", string.Empty);
        if (!string.IsNullOrWhiteSpace(style))
        {
            var cleaned = RemoveBorderStyles(style);
            table.SetAttributeValue("style", cleaned);
        }

        AppendInlineStyleIfMissing(table, "border:none; border-collapse:collapse; border-spacing:0;", "border");
    }

    private static string RemoveBorderStyles(string styleValue)
    {
        var parts = styleValue.Split(';', StringSplitOptions.RemoveEmptyEntries)
            .Select(part => part.Trim());

        var keptParts = parts.Where(part =>
            !part.StartsWith("border", StringComparison.OrdinalIgnoreCase)
            && !part.StartsWith("border-", StringComparison.OrdinalIgnoreCase)
            && !part.StartsWith("border-collapse", StringComparison.OrdinalIgnoreCase)
            && !part.StartsWith("border-spacing", StringComparison.OrdinalIgnoreCase));

        return string.Join("; ", keptParts);
    }

    private static void AppendInlineStyleIfMissing(HtmlNode node, string styleFragment, string propertyName)
    {
        var existing = node.GetAttributeValue("style", string.Empty);
        if (existing.Contains(propertyName, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        var merged = string.IsNullOrWhiteSpace(existing)
            ? styleFragment
            : $"{existing.TrimEnd(';')}; {styleFragment}";

        node.SetAttributeValue("style", merged);
    }
}
