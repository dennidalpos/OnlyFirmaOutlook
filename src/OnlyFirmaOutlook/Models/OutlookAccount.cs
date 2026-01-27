namespace OnlyFirmaOutlook.Models;




public class OutlookAccount
{
    public string DisplayName { get; set; } = string.Empty;
    public string SmtpAddress { get; set; } = string.Empty;
    public string AccountType { get; set; } = string.Empty;
    public string StoreId { get; set; } = string.Empty;
    public string OwnerSmtpAddress { get; set; } = string.Empty;
    public bool IsSharedMailbox { get; set; }

    public string DisplayText
    {
        get
        {
            var baseText = !string.IsNullOrEmpty(SmtpAddress)
                ? SmtpAddress
                : DisplayName;

            if (IsSharedMailbox)
            {
                return string.IsNullOrEmpty(baseText)
                    ? $"{DisplayName} (shared)"
                    : $"{baseText} (shared)";
            }

            return baseText;
        }
    }

    public override string ToString() => DisplayText;
}
