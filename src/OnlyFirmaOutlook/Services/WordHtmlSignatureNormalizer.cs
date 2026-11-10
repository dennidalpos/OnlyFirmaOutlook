// (c) 2026 Danny Perondi. All rights reserved. Proprietary and confidential.

using System;
using HtmlAgilityPack;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;

namespace OnlyFirmaOutlook.Services;

/// <summary>Normalizza HTML Word per firme Outlook preservando formattazioni visive.</summary>
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

        RemoveNonRenderingElements(workingDoc);
        CleanupStyles(workingDoc);

        return workingDoc.DocumentNode.InnerHtml;
    }

    /// <summary>Rimuove elementi non visibili/non supportati (script, meta, xml, o:p, commenti, w:*).</summary>
    private static void RemoveNonRenderingElements(HtmlDocument doc)
    {
        var nodesToRemove = doc.DocumentNode
            .Descendants()
            .Where(node => HasNodeName(node, "script", "meta", "xml", "o:p"))
            .ToList();

        if (nodesToRemove.Count > 0)
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

        var wordNamespaceNodes = doc.DocumentNode.Descendants()
            .Where(node => node.Name.StartsWith("w:", StringComparison.OrdinalIgnoreCase))
            .ToList();

        foreach (var node in wordNamespaceNodes)
        {
            node.Remove();
        }
    }

    /// <summary>Rimuove stili Office non supportati mantenendo quelli visivi standard.</summary>
    private static void CleanupStyles(HtmlDocument doc)
    {
        foreach (var node in doc.DocumentNode.Descendants())
        {
            var style = node.GetAttributeValue("style", null);
            if (style == null)
            {
                continue;
            }

            var cleaned = RemoveUnsupportedStyles(style);
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

    /// <summary>Filtra proprietà mso-* (eccetto mso-line-height-rule) e tab-stops.</summary>
    private static string RemoveUnsupportedStyles(string styleValue)
    {
        var parts = styleValue.Split(';', StringSplitOptions.RemoveEmptyEntries)
            .Select(part => part.Trim())
            .Where(part =>
            {
                // Filtra mso-* tranne mso-line-height-rule
                if (part.StartsWith("mso-", StringComparison.OrdinalIgnoreCase))
                {
                    return part.StartsWith("mso-line-height-rule", StringComparison.OrdinalIgnoreCase);
                }

                // Filtra tab-stops
                if (part.Contains("tab-stops", StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                return true;
            });

        return string.Join("; ", parts);
    }

    private static bool HasNodeName(HtmlNode node, params string[] nodeNames)
    {
        return nodeNames.Any(nodeName => string.Equals(node.Name, nodeName, StringComparison.OrdinalIgnoreCase));
    }
}
