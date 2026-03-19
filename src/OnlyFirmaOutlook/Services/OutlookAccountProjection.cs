/*
 * OnlyFirmaOutlook
 * Copyright (c) 2026 Danny Perondi. All rights reserved.
 * Author: Danny Perondi
 * Proprietary and confidential.
 * Unauthorized copying, modification, distribution, sublicensing, disclosure,
 * or commercial use is prohibited without prior written permission.
 */

using OnlyFirmaOutlook.Models;

namespace OnlyFirmaOutlook.Services;

internal static class OutlookAccountProjection
{
    internal sealed record DelegatedStoreCandidate(string DisplayName, bool IsDelegatedStore);

    internal static OutlookAccount CreateAccount(string? displayName, string? smtpAddress, int accountType)
    {
        return new OutlookAccount
        {
            DisplayName = string.IsNullOrWhiteSpace(displayName) ? "Account senza nome" : displayName,
            SmtpAddress = smtpAddress ?? string.Empty,
            AccountType = MapAccountType(accountType)
        };
    }

    internal static List<OutlookAccount> MergeAccountsWithDelegates(
        IEnumerable<OutlookAccount> accounts,
        IEnumerable<DelegatedStoreCandidate> stores)
    {
        var mergedAccounts = accounts.ToList();
        var accountIdentifiers = BuildAccountIdentifiers(mergedAccounts);

        foreach (var store in stores)
        {
            if (string.IsNullOrWhiteSpace(store.DisplayName))
            {
                continue;
            }

            if (!store.IsDelegatedStore)
            {
                continue;
            }

            if (accountIdentifiers.Contains(store.DisplayName))
            {
                continue;
            }

            mergedAccounts.Add(new OutlookAccount
            {
                DisplayName = store.DisplayName,
                SmtpAddress = string.Empty,
                AccountType = "Delega",
                IsDelegate = true
            });

            accountIdentifiers.Add(store.DisplayName);
        }

        return SortAccounts(mergedAccounts);
    }

    internal static List<OutlookAccount> SortAccounts(IEnumerable<OutlookAccount> accounts)
    {
        return accounts
            .OrderBy(account => account.IsDelegate)
            .ThenBy(account => account.DisplayText)
            .ToList();
    }

    internal static string MapAccountType(int accountType)
    {
        return accountType switch
        {
            1 => "Exchange",
            2 => "IMAP",
            3 => "POP3",
            4 => "HTTP",
            5 => "EAS",
            _ => "Altro"
        };
    }

    private static HashSet<string> BuildAccountIdentifiers(IEnumerable<OutlookAccount> accounts)
    {
        var identifiers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var account in accounts)
        {
            if (!string.IsNullOrWhiteSpace(account.DisplayName))
            {
                identifiers.Add(account.DisplayName);
            }

            if (!string.IsNullOrWhiteSpace(account.SmtpAddress))
            {
                identifiers.Add(account.SmtpAddress);
            }
        }

        return identifiers;
    }
}
