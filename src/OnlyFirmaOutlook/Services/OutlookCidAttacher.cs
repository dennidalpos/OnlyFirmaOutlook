using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;
using OnlyFirmaOutlook.Models;

namespace OnlyFirmaOutlook.Services;

public static class OutlookCidAttacher
{
    private const string PrAttachContentId = "http://schemas.microsoft.com/mapi/proptag/0x3712001F";
    private const string PrAttachContentLocation = "http://schemas.microsoft.com/mapi/proptag/0x3713001F";
    private const string PrAttachContentDisposition = "http://schemas.microsoft.com/mapi/proptag/0x3716001F";
    private const string PrAttachmentHidden = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B";

    public static void AddInlineCidAttachments(MailItem mailItem, IEnumerable<InlineImage> images)
    {
        if (mailItem is null)
        {
            throw new ArgumentNullException(nameof(mailItem));
        }

        if (images is null)
        {
            throw new ArgumentNullException(nameof(images));
        }

        Attachments? attachments = null;

        try
        {
            attachments = mailItem.Attachments;
            foreach (var image in images)
            {
                if (string.IsNullOrWhiteSpace(image.FilePath) || !File.Exists(image.FilePath))
                {
                    continue;
                }

                Attachment? attachment = null;
                PropertyAccessor? accessor = null;

                try
                {
                    attachment = attachments.Add(
                        image.FilePath,
                        OlAttachmentType.olByValue,
                        Type.Missing,
                        image.FileName);

                    accessor = attachment.PropertyAccessor;
                    accessor.SetProperty(PrAttachContentId, image.ContentId);
                    accessor.SetProperty(PrAttachContentLocation, image.FileName);
                    accessor.SetProperty(PrAttachContentDisposition, "inline");
                    accessor.SetProperty(PrAttachmentHidden, true);
                }
                finally
                {
                    if (accessor != null)
                    {
                        Marshal.FinalReleaseComObject(accessor);
                    }

                    if (attachment != null)
                    {
                        Marshal.FinalReleaseComObject(attachment);
                    }
                }
            }
        }
        finally
        {
            if (attachments != null)
            {
                Marshal.FinalReleaseComObject(attachments);
            }
        }
    }
}
