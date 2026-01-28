using System;
using System.Collections.Generic;
using HtmlAgilityPack;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;

namespace OnlyFirmaOutlook.Services;

public class WordHtmlSignatureNormalizer
{
    public string Normalize(string html)
    {
        var correctedHtml = ApplyManualBorderFixes(html);
        var doc = new HtmlDocument();
        doc.LoadHtml(correctedHtml);

        var bodyNode = doc.DocumentNode.SelectSingleNode("//body");
        var workingHtml = bodyNode?.InnerHtml ?? doc.DocumentNode.InnerHtml;

        var workingDoc = new HtmlDocument();
        workingDoc.LoadHtml(workingHtml);

        RemoveUnsafeNodes(workingDoc);
        NormalizeNodes(workingDoc);
        NormalizeTableGridClasses(workingDoc);
        NormalizeTableBorders(workingDoc);
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
            RemoveTableBorders(table.Descendants());
        }
    }

    private static void NormalizeTableGridClasses(HtmlDocument doc)
    {
        var tableNodes = doc.DocumentNode.SelectNodes("//table");
        if (tableNodes == null)
        {
            return;
        }

        foreach (var table in tableNodes)
        {
            var classValue = table.GetAttributeValue("class", string.Empty);
            if (!classValue.Contains("MsoTableGrid", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var cleaned = string.Join(' ', classValue
                .Split(' ', StringSplitOptions.RemoveEmptyEntries)
                .Where(value => !value.Equals("MsoTableGrid", StringComparison.OrdinalIgnoreCase)));

            if (string.IsNullOrWhiteSpace(cleaned))
            {
                table.Attributes.Remove("class");
            }
            else
            {
                table.SetAttributeValue("class", cleaned);
            }
        }
    }

    private static void NormalizeTableBorders(HtmlDocument doc)
    {
        var tableNodes = doc.DocumentNode.SelectNodes("//table");
        if (tableNodes == null)
        {
            return;
        }

        foreach (var table in tableNodes)
        {
            if (!HasBorderIndicators(table))
            {
                var hasBorderInCells = table
                    .Descendants()
                    .Any(node => node.Name is "tr" or "td" or "th" && HasBorderIndicators(node));

                if (!hasBorderInCells)
                {
                    continue;
                }
            }

            ApplyBorderReset(table);
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

    private static string ApplyManualBorderFixes(string html)
    {
        if (string.IsNullOrWhiteSpace(html))
        {
            return html;
        }

        return html
            .Replace("mso-border-alt:solid", "mso-border-alt:none", StringComparison.OrdinalIgnoreCase)
            .Replace("border:solid", "border:none", StringComparison.OrdinalIgnoreCase)
            .Replace("border-style:solid", "border-style:none", StringComparison.OrdinalIgnoreCase);
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
        RemoveTableBorderAttributes(table);

        var style = table.GetAttributeValue("style", string.Empty);
        if (!string.IsNullOrWhiteSpace(style))
        {
            var cleaned = RemoveBorderStyles(style);
            table.SetAttributeValue("style", cleaned);
        }

        AppendInlineStyleIfMissing(table, "border:none; border-collapse:collapse; border-spacing:0;", "border");
    }

    private static void ApplyBorderReset(HtmlNode table)
    {
        RemoveTableBorderAttributes(table);
        table.SetAttributeValue("border", "0");
        table.SetAttributeValue("cellspacing", "0");
        table.SetAttributeValue("cellpadding", "0");

        var style = table.GetAttributeValue("style", string.Empty);
        if (!string.IsNullOrWhiteSpace(style))
        {
            var cleaned = RemoveBorderStyles(style);
            if (string.IsNullOrWhiteSpace(cleaned))
            {
                table.Attributes.Remove("style");
            }
            else
            {
                table.SetAttributeValue("style", cleaned);
            }
        }

        AppendInlineStyleIfMissing(table, "border:none; border-collapse:collapse; border-spacing:0;", "border");

        foreach (var node in table.Descendants().Where(n => n.Name is "table" or "tr" or "td" or "th"))
        {
            RemoveTableBorderAttributes(node);

            var nodeStyle = node.GetAttributeValue("style", string.Empty);
            if (!string.IsNullOrWhiteSpace(nodeStyle))
            {
                var cleaned = RemoveBorderStyles(nodeStyle);
                if (string.IsNullOrWhiteSpace(cleaned))
                {
                    node.Attributes.Remove("style");
                }
                else
                {
                    node.SetAttributeValue("style", cleaned);
                }
            }

            AppendInlineStyleIfMissing(node, "border:none;", "border");
        }
    }

    private static void RemoveTableBorders(IEnumerable<HtmlNode> nodes)
    {
        foreach (var node in nodes)
        {
            if (node.Name is not ("table" or "tr" or "td" or "th"))
            {
                continue;
            }

            RemoveTableBorderAttributes(node);

            var style = node.GetAttributeValue("style", string.Empty);
            if (!string.IsNullOrWhiteSpace(style))
            {
                var cleaned = RemoveBorderStyles(style);
                if (string.IsNullOrWhiteSpace(cleaned))
                {
                    node.Attributes.Remove("style");
                }
                else
                {
                    node.SetAttributeValue("style", cleaned);
                }
            }
        }
    }

    private static void RemoveTableBorderAttributes(HtmlNode node)
    {
        node.Attributes.Remove("border");
        node.Attributes.Remove("cellspacing");
        node.Attributes.Remove("cellpadding");
        node.Attributes.Remove("bordercolor");
        node.Attributes.Remove("bordercolorlight");
        node.Attributes.Remove("bordercolordark");
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

    private static bool HasBorderIndicators(HtmlNode node)
    {
        var classValue = node.GetAttributeValue("class", string.Empty);
        if (classValue.Contains("MsoTableGrid", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        var style = node.GetAttributeValue("style", string.Empty);
        if (style.Contains("mso-border", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return style.Contains("border:solid", StringComparison.OrdinalIgnoreCase)
               || style.Contains("border-style:solid", StringComparison.OrdinalIgnoreCase);
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
