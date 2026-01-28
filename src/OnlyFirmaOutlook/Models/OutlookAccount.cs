namespace OnlyFirmaOutlook.Models;




public class OutlookAccount
{
    public string DisplayName { get; set; } = string.Empty;
    public string SmtpAddress { get; set; } = string.Empty;
    public string AccountType { get; set; } = string.Empty;
    public bool IsDelegate { get; set; }

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

    public string GroupLabel => IsDelegate ? "Deleghe" : "Account";

    public override string ToString() => DisplayText;
}
