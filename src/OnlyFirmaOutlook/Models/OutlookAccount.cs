/*
 * OnlyFirmaOutlook
 * Copyright (c) 2026 Danny Perondi. All rights reserved.
 * Author: Danny Perondi
 * Proprietary and confidential.
 * Unauthorized copying, modification, distribution, sublicensing, disclosure,
 * or commercial use is prohibited without prior written permission.
 */

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
