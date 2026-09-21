// (c) 2026 Danny Perondi. All rights reserved. Proprietary and confidential.

namespace OnlyFirmaOutlook.Models;

public class OutlookAccount
{
    public string DisplayName { get; set; } = string.Empty;
    public string SmtpAddress { get; set; } = string.Empty;
    public string AccountType { get; set; } = string.Empty;
    public bool IsDelegate { get; set; }

    public string DisplayText => !string.IsNullOrEmpty(SmtpAddress) ? SmtpAddress : DisplayName;

    public string GroupLabel => IsDelegate ? "Deleghe" : "Account";

    public override string ToString() => DisplayText;
}
