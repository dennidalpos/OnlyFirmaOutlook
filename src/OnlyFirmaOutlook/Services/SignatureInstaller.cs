using System.IO;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;

namespace OnlyFirmaOutlook.Services;

public class SignatureInstaller
{
    private readonly LoggingService _logger;

    public SignatureInstaller()
    {
        _logger = LoggingService.Instance;
    }

    public void Install(string destinationFolder, string signatureName, string html, string plainText)
    {
        Directory.CreateDirectory(destinationFolder);
        var htmlPath = Path.Combine(destinationFolder, signatureName + ".htm");
        var txtPath = Path.Combine(destinationFolder, signatureName + ".txt");

        File.WriteAllText(htmlPath, html);
        File.WriteAllText(txtPath, plainText);

        _logger.Log("Firma HTML e TXT aggiornate");
    }

    public static string BuildPlainText(string html)
    {
        var doc = new HtmlDocument();
        doc.LoadHtml(html);

        var nodes = doc.DocumentNode.SelectNodes("//br|//p|//div|//tr");
        if (nodes != null)
        {
            foreach (var node in nodes)
            {
                node.AppendChild(doc.CreateTextNode("\n"));
            }
        }

        var text = doc.DocumentNode.InnerText;
        text = Regex.Replace(text, @"\r\n|\r", "\n");
        text = Regex.Replace(text, @"\n{3,}", "\n\n");
        text = Regex.Replace(text, @"[ \t]+", " ");
        return text.Trim();
    }
}
