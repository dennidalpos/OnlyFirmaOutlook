using System;
using HtmlAgilityPack;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;

namespace OnlyFirmaOutlook.Services;

/// <summary>
/// Normalizza HTML generato da Word per firme Outlook.
/// Preserva formattazioni volute: font, dimensioni, spaziature, elenchi puntati.
/// </summary>
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

    /// <summary>
    /// Rimuove elementi che non hanno impatto visivo e non sono supportati dai client email.
    /// </summary>
    private static void RemoveNonRenderingElements(HtmlDocument doc)
    {
        // Script e meta: sicurezza e non supportati
        // xml: markup Office non renderizzato
        // o:p: paragrafo vuoto Office (solo spaziatura artificiale)
        var nodesToRemove = doc.DocumentNode.SelectNodes("//script|//meta|//xml|//o:p");
        if (nodesToRemove != null)
        {
            foreach (var node in nodesToRemove)
            {
                node.Remove();
            }
        }

        // Commenti HTML: non visibili, appesantiscono
        var commentNodes = doc.DocumentNode.SelectNodes("//comment()");
        if (commentNodes != null)
        {
            foreach (var comment in commentNodes)
            {
                comment.Remove();
            }
        }

        // Tag Word namespace (w:*): non supportati da client email
        var wordNamespaceNodes = doc.DocumentNode.Descendants()
            .Where(node => node.Name.StartsWith("w:", StringComparison.OrdinalIgnoreCase))
            .ToList();

        foreach (var node in wordNamespaceNodes)
        {
            node.Remove();
        }
    }

    /// <summary>
    /// Rimuove solo stili Microsoft Office che non sono supportati universalmente
    /// e non hanno impatto visivo. Preserva tutto il resto.
    /// </summary>
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
                node.Attributes["style"].Value = cleaned;
            }
        }
    }

    /// <summary>
    /// Rimuove stili mso-* che non sono supportati da client email standard.
    /// Preserva: font-family, font-size, color, margin, padding, line-height, text-align, etc.
    /// </summary>
    private static string RemoveUnsupportedStyles(string styleValue)
    {
        var parts = styleValue.Split(';', StringSplitOptions.RemoveEmptyEntries)
            .Select(part => part.Trim())
            .Where(part =>
            {
                // Rimuove solo proprietà mso-* (Microsoft Office specific)
                // ECCEZIONE: mso-line-height-rule è utile per Outlook
                if (part.StartsWith("mso-", StringComparison.OrdinalIgnoreCase))
                {
                    return part.StartsWith("mso-line-height-rule", StringComparison.OrdinalIgnoreCase);
                }

                // Rimuove tab-stops (Word specific, non supportato)
                if (part.Contains("tab-stops", StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                // Preserva tutto il resto
                return true;
            });

        return string.Join("; ", parts);
    }

}
