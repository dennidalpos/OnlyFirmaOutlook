using OnlyFirmaOutlook.Models;
using OnlyFirmaOutlook.Services;

namespace OnlyFirmaOutlook.Tests;

public class OutlookAccountProjectionTests
{
    [Theory]
    [InlineData(1, "Exchange")]
    [InlineData(2, "IMAP")]
    [InlineData(3, "POP3")]
    [InlineData(4, "HTTP")]
    [InlineData(5, "EAS")]
    [InlineData(99, "Altro")]
    public void MapAccountType_ReturnsExpectedLabel(int accountType, string expected)
    {
        var result = OutlookAccountProjection.MapAccountType(accountType);

        Assert.Equal(expected, result);
    }

    [Fact]
    public void CreateAccount_NormalizesMissingDisplayName()
    {
        var account = OutlookAccountProjection.CreateAccount("   ", null, 2);

        Assert.Equal("Account senza nome", account.DisplayName);
        Assert.Equal(string.Empty, account.SmtpAddress);
        Assert.Equal("IMAP", account.AccountType);
    }

    [Fact]
    public void MergeAccountsWithDelegates_AddsOnlyUniqueDelegatedStores()
    {
        var accounts = new[]
        {
            new OutlookAccount
            {
                DisplayName = "Mario Rossi",
                SmtpAddress = "mario.rossi@example.com",
                AccountType = "Exchange"
            }
        };

        var stores = new[]
        {
            new OutlookAccountProjection.DelegatedStoreCandidate("Mario Rossi", true),
            new OutlookAccountProjection.DelegatedStoreCandidate("mario.rossi@example.com", true),
            new OutlookAccountProjection.DelegatedStoreCandidate("Shared Mailbox", false),
            new OutlookAccountProjection.DelegatedStoreCandidate("Support Team", true),
            new OutlookAccountProjection.DelegatedStoreCandidate("Support Team", true),
            new OutlookAccountProjection.DelegatedStoreCandidate("", true)
        };

        var merged = OutlookAccountProjection.MergeAccountsWithDelegates(accounts, stores);

        Assert.Equal(2, merged.Count);
        Assert.Contains(merged, account => !account.IsDelegate && account.DisplayText == "mario.rossi@example.com");
        Assert.Contains(merged, account => account.IsDelegate && account.DisplayName == "Support Team");
    }

    [Fact]
    public void SortAccounts_KeepsPrimaryAccountsBeforeDelegatesAndOrdersByDisplayText()
    {
        var accounts = new[]
        {
            new OutlookAccount { DisplayName = "Shared B", AccountType = "Delega", IsDelegate = true },
            new OutlookAccount { DisplayName = "Zeta", SmtpAddress = "zeta@example.com", AccountType = "Exchange" },
            new OutlookAccount { DisplayName = "Shared A", AccountType = "Delega", IsDelegate = true },
            new OutlookAccount { DisplayName = "Alfa", SmtpAddress = "alfa@example.com", AccountType = "IMAP" }
        };

        var sorted = OutlookAccountProjection.SortAccounts(accounts);

        Assert.Collection(
            sorted,
            account => Assert.Equal("alfa@example.com", account.DisplayText),
            account => Assert.Equal("zeta@example.com", account.DisplayText),
            account => Assert.Equal("Shared A", account.DisplayText),
            account => Assert.Equal("Shared B", account.DisplayText));
    }
}
