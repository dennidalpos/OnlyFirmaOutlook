using System.IO;
using System.Text;
using Microsoft.Office.Interop.Outlook;

namespace OnlyFirmaOutlook.Services;

public static class OutlookSignatureEmbedder
{
    public static void ApplySignatureWithInlineImages(MailItem mailItem, string htmlPath)
    {
        if (mailItem is null)
        {
            throw new ArgumentNullException(nameof(mailItem));
        }

        if (string.IsNullOrWhiteSpace(htmlPath))
        {
            throw new ArgumentException("HTML path is required.", nameof(htmlPath));
        }

        if (!File.Exists(htmlPath))
        {
            throw new FileNotFoundException("Signature HTML file not found.", htmlPath);
        }

        var html = File.ReadAllText(htmlPath, Encoding.GetEncoding(1252));
        var baseDir = Path.GetDirectoryName(htmlPath) ?? string.Empty;
        var (htmlRewritten, images) = WordHtmlCidPostProcessor.RewriteLocalImageRefsToCid(html, baseDir);

        mailItem.BodyFormat = OlBodyFormat.olFormatHTML;
        mailItem.HTMLBody = htmlRewritten;
        OutlookCidAttacher.AddInlineCidAttachments(mailItem, images);
    }
}
