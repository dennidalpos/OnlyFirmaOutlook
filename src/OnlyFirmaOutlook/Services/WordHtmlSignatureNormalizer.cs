using System;
using HtmlAgilityPack;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;

namespace OnlyFirmaOutlook.Services;

/// <summary>
/// Normalizza HTML generato da Word per firme Outlook.
/// Risolve il bug di Outlook 2512+ che aggiunge bordi alle tabelle.
/// Preserva formattazioni volute: font, dimensioni, spaziature, elenchi puntati.
/// </summary>
public class WordHtmlSignatureNormalizer
{
    public string Normalize(string html, bool fixOutlook2512 = true)
    {
        var doc = new HtmlDocument();
        doc.LoadHtml(html);

        var bodyNode = doc.DocumentNode.SelectSingleNode("//body");
        var workingHtml = bodyNode?.InnerHtml ?? doc.DocumentNode.InnerHtml;

        var workingDoc = new HtmlDocument();
        workingDoc.LoadHtml(workingHtml);

        RemoveNonRenderingElements(workingDoc);
        CleanupStyles(workingDoc);

        if (fixOutlook2512)
        {
            FixTableBordersForOutlook2512(workingDoc);
        }

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

    /// <summary>
    /// Fix per bug Outlook Classic 2512+: aggiunge bordi indesiderati alle tabelle.
    /// Imposta esplicitamente border:none su tabelle e celle.
    /// </summary>
    private static void FixTableBordersForOutlook2512(HtmlDocument doc)
    {
        var tables = doc.DocumentNode.SelectNodes("//table");
        if (tables != null)
        {
            foreach (var table in tables)
            {
                FixTableBorders(table);
            }
        }

        var cells = doc.DocumentNode.SelectNodes("//td|//th");
        if (cells != null)
        {
            foreach (var cell in cells)
            {
                FixCellBorders(cell);
            }
        }
    }

    private static void FixTableBorders(HtmlNode tableNode)
    {
        // Attributi HTML per compatibilità
        tableNode.SetAttributeValue("border", "0");
        tableNode.SetAttributeValue("cellpadding", "0");
        tableNode.SetAttributeValue("cellspacing", "0");

        // Rimuove bordi esistenti dallo style e aggiunge fix
        var existingStyle = tableNode.GetAttributeValue("style", string.Empty);
        var cleanedStyle = RemoveExistingBorderStyles(existingStyle);

        // Stili necessari per fix Outlook 2512
        const string borderFix = "border:none; border-collapse:collapse";

        var newStyle = string.IsNullOrWhiteSpace(cleanedStyle)
            ? borderFix
            : $"{cleanedStyle.TrimEnd(';')}; {borderFix}";

        tableNode.SetAttributeValue("style", newStyle);
    }

    private static void FixCellBorders(HtmlNode cellNode)
    {
        var existingStyle = cellNode.GetAttributeValue("style", string.Empty);
        var cleanedStyle = RemoveExistingBorderStyles(existingStyle);

        // Aggiunge border:none solo se non già presente
        if (!cleanedStyle.Contains("border", StringComparison.OrdinalIgnoreCase))
        {
            var newStyle = string.IsNullOrWhiteSpace(cleanedStyle)
                ? "border:none"
                : $"{cleanedStyle.TrimEnd(';')}; border:none";

            cellNode.SetAttributeValue("style", newStyle);
        }
        else if (!string.IsNullOrWhiteSpace(cleanedStyle))
        {
            cellNode.SetAttributeValue("style", cleanedStyle);
        }
    }

    /// <summary>
    /// Rimuove stili di bordo problematici inseriti da Word/Outlook 2512.
    /// </summary>
    private static string RemoveExistingBorderStyles(string styleValue)
    {
        if (string.IsNullOrWhiteSpace(styleValue))
        {
            return string.Empty;
        }

        var parts = styleValue.Split(';', StringSplitOptions.RemoveEmptyEntries)
            .Select(part => part.Trim())
            .Where(part =>
            {
                // Rimuove tutti gli stili border (verranno reimpostati)
                if (part.StartsWith("border", StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                // Rimuove windowtext (colore bordo Word)
                if (part.Contains("windowtext", StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                return true;
            });

        return string.Join("; ", parts);
    }
}
