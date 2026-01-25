namespace OnlyFirmaOutlook.Models;

/// <summary>
/// Rappresenta un account Outlook.
/// </summary>
public class OutlookAccount
{
    public string DisplayName { get; set; } = string.Empty;
    public string SmtpAddress { get; set; } = string.Empty;
    public string AccountType { get; set; } = string.Empty;

    public string DisplayText
    {
        get
        {
            if (!string.IsNullOrEmpty(SmtpAddress))
            {
                return SmtpAddress;
            }
            return DisplayName;
        }
    }

    public override string ToString() => DisplayText;
}
