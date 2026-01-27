using System.Runtime.InteropServices;
using OnlyFirmaOutlook.Models;

namespace OnlyFirmaOutlook.Services;





public class OutlookAccountService
{
    private readonly LoggingService _logger;

    public OutlookAccountService()
    {
        _logger = LoggingService.Instance;
    }

    
    
    
    public class AccountLoadResult
    {
        public bool OutlookAvailable { get; set; }
        public List<OutlookAccount> Accounts { get; set; } = new();
        public string? ErrorMessage { get; set; }
    }

    
    
    
    
    public AccountLoadResult LoadAccounts()
    {
        _logger.Log("Caricamento account Outlook...");

        var result = new AccountLoadResult();
        dynamic? outlookApp = null;
        dynamic? session = null;
        dynamic? accounts = null;
        dynamic? stores = null;

        try
        {
            
            var outlookType = Type.GetTypeFromProgID("Outlook.Application");
            if (outlookType == null)
            {
                _logger.Log("Outlook non installato");
                result.OutlookAvailable = false;
                return result;
            }

            
            _logger.Log("Creazione istanza Outlook.Application...");
            outlookApp = Activator.CreateInstance(outlookType);
            if (outlookApp == null)
            {
                result.OutlookAvailable = false;
                result.ErrorMessage = "Impossibile creare istanza di Outlook";
                _logger.LogError(result.ErrorMessage);
                return result;
            }

            result.OutlookAvailable = true;

            
            _logger.Log("Accesso alla sessione MAPI...");
            session = outlookApp.Session;
            if (session == null)
            {
                result.ErrorMessage = "Impossibile accedere alla sessione Outlook";
                _logger.LogWarning(result.ErrorMessage);
                return result;
            }

            
            accounts = session.Accounts;
            if (accounts == null)
            {
                _logger.Log("Nessun account Outlook configurato");
                return result;
            }

            int accountCount = accounts.Count;
            if (accountCount == 0)
            {
                _logger.Log("Nessun account Outlook configurato");
                return result;
            }

            _logger.Log($"Trovati {accountCount} account");

            
            for (int i = 1; i <= accountCount; i++)
            {
                dynamic? account = null;
                try
                {
                    account = accounts.Item(i);
                    if (account != null)
                    {
                        dynamic? accountStore = account.Store;
                        var storeId = accountStore?.StoreID as string;
                        var isDelegated = IsDelegatedAccount(account);
                        var outlookAccount = new OutlookAccount
                        {
                            DisplayName = account.DisplayName ?? "Account senza nome",
                            SmtpAddress = account.SmtpAddress ?? string.Empty,
                            AccountType = isDelegated ? "Delegato" : GetAccountTypeName(account.AccountType),
                            IsDelegated = isDelegated,
                            StoreId = storeId
                        };

                        result.Accounts.Add(outlookAccount);
                        _logger.Log($"Account trovato: {outlookAccount.DisplayText} ({outlookAccount.AccountType})");
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning($"Errore lettura account {i}: {ex.Message}");
                }
                finally
                {
                    if (account != null)
                    {
                        try { Marshal.FinalReleaseComObject(account); }
                        catch {  }
                    }
                }
            }

            _logger.Log($"Totale account caricati: {result.Accounts.Count}");

            var accountStoreIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var accountSmtpAddresses = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var accountDisplayNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                for (int i = 1; i <= accountCount; i++)
                {
                    dynamic? account = null;
                    try
                    {
                        account = accounts.Item(i);
                        if (account != null)
                        {
                            dynamic? accountStore = account.Store;
                            if (accountStore != null)
                            {
                                var storeId = accountStore.StoreID as string;
                                if (!string.IsNullOrEmpty(storeId))
                                {
                                    accountStoreIds.Add(storeId);
                                }
                            }

                            var smtp = account.SmtpAddress as string;
                            if (!string.IsNullOrWhiteSpace(smtp))
                            {
                                accountSmtpAddresses.Add(smtp);
                            }

                            var displayName = account.DisplayName as string;
                            if (!string.IsNullOrWhiteSpace(displayName))
                            {
                                accountDisplayNames.Add(displayName);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning($"Errore lettura store account {i}: {ex.Message}");
                    }
                    finally
                    {
                        if (account != null)
                        {
                            try { Marshal.FinalReleaseComObject(account); }
                            catch { }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Errore raccolta store account: {ex.Message}");
            }

            try
            {
                stores = session.Stores;
                if (stores != null)
                {
                    var storeCount = stores.Count;
                    for (int i = 1; i <= storeCount; i++)
                    {
                        dynamic? store = null;
                        try
                        {
                            store = stores.Item(i);
                            if (store == null)
                            {
                                continue;
                            }

                            var storeId = store.StoreID as string;
                            if (!string.IsNullOrEmpty(storeId) && accountStoreIds.Contains(storeId))
                            {
                                continue;
                            }

                            var isDataFileStore = false;
                            try
                            {
                                isDataFileStore = store.IsDataFileStore;
                            }
                            catch
                            {
                                isDataFileStore = false;
                            }

                            if (isDataFileStore)
                            {
                                continue;
                            }

                            var displayName = store.DisplayName as string ?? "Mailbox delegata";
                            if (IsOnlineArchiveStore(store, displayName))
                            {
                                continue;
                            }

                            var smtpAddress = TryGetStoreSmtpAddress(store);
                            if (IsDuplicateStore(displayName, smtpAddress, accountSmtpAddresses, accountDisplayNames))
                            {
                                continue;
                            }

                            var delegatedAccount = new OutlookAccount
                            {
                                DisplayName = displayName,
                                SmtpAddress = smtpAddress ?? string.Empty,
                                AccountType = "Delegato",
                                IsDelegated = true,
                                StoreId = storeId
                            };

                            result.Accounts.Add(delegatedAccount);
                            _logger.Log($"Mailbox delegata trovata: {delegatedAccount.DisplayText}");
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning($"Errore lettura store {i}: {ex.Message}");
                        }
                        finally
                        {
                            if (store != null)
                            {
                                try { Marshal.FinalReleaseComObject(store); }
                                catch { }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Errore accesso store Outlook: {ex.Message}");
            }
        }
        catch (COMException comEx)
        {
            result.ErrorMessage = $"Errore COM Outlook: {comEx.Message}";
            _logger.LogError(result.ErrorMessage, comEx);

            
            if (comEx.ErrorCode == unchecked((int)0x80040154) || 
                comEx.ErrorCode == unchecked((int)0x80080005))   
            {
                result.OutlookAvailable = false;
            }
        }
        catch (Exception ex)
        {
            result.ErrorMessage = $"Errore durante il caricamento account: {ex.Message}";
            _logger.LogError(result.ErrorMessage, ex);
        }
        finally
        {
            
            CleanupComObjects(stores, accounts, session, outlookApp);
        }

        return result;
    }

    
    
    
    private void CleanupComObjects(dynamic? stores, dynamic? accounts, dynamic? session, dynamic? outlookApp)
    {
        _logger.Log("Cleanup oggetti COM Outlook...");

        try
        {
            if (stores != null)
            {
                try { Marshal.FinalReleaseComObject(stores); }
                catch (Exception ex)
                {
                    _logger.LogWarning($"Errore rilascio stores COM: {ex.Message}");
                }
            }
        }
        catch { }

        try
        {
            if (accounts != null)
            {
                try { Marshal.FinalReleaseComObject(accounts); }
                catch (Exception ex)
                {
                    _logger.LogWarning($"Errore rilascio accounts COM: {ex.Message}");
                }
            }
        }
        catch {  }

        try
        {
            if (session != null)
            {
                try { Marshal.FinalReleaseComObject(session); }
                catch (Exception ex)
                {
                    _logger.LogWarning($"Errore rilascio session COM: {ex.Message}");
                }
            }
        }
        catch {  }

        try
        {
            if (outlookApp != null)
            {
                
                try { Marshal.FinalReleaseComObject(outlookApp); }
                catch (Exception ex)
                {
                    _logger.LogWarning($"Errore rilascio Outlook COM: {ex.Message}");
                }
            }
        }
        catch {  }

        GC.Collect();
        GC.WaitForPendingFinalizers();

        _logger.Log("Cleanup COM Outlook completato");
    }

    private static bool IsDelegatedAccount(dynamic account)
    {
        try
        {
            dynamic? deliveryStore = account.DeliveryStore;
            var deliveryStoreId = deliveryStore?.StoreID as string;

            if (deliveryStore == null)
            {
                return true;
            }

            if (string.IsNullOrWhiteSpace(deliveryStoreId))
            {
                return true;
            }
        }
        catch
        {
            return false;
        }

        return false;
    }

    private static bool IsOnlineArchiveStore(dynamic store, string displayName)
    {
        try
        {
            var isArchive = store.IsArchive;
            if (isArchive is bool boolValue && boolValue)
            {
                return true;
            }
        }
        catch
        {
        }

        if (displayName.Contains("Archivio online", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (displayName.Contains("Online Archive", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return false;
    }

    private static bool IsDuplicateStore(
        string displayName,
        string? smtpAddress,
        HashSet<string> accountSmtpAddresses,
        HashSet<string> accountDisplayNames)
    {
        if (!string.IsNullOrWhiteSpace(smtpAddress) && accountSmtpAddresses.Contains(smtpAddress))
        {
            return true;
        }

        if (accountDisplayNames.Contains(displayName))
        {
            return true;
        }

        foreach (var accountSmtp in accountSmtpAddresses)
        {
            if (!string.IsNullOrWhiteSpace(accountSmtp) &&
                displayName.Contains(accountSmtp, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        if (!string.IsNullOrWhiteSpace(smtpAddress))
        {
            foreach (var accountName in accountDisplayNames)
            {
                if (!string.IsNullOrWhiteSpace(accountName) &&
                    smtpAddress.Contains(accountName, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
        }

        return false;
    }

    private string? TryGetStoreSmtpAddress(dynamic store)
    {
        try
        {
            var propertyAccessor = store.PropertyAccessor;
            if (propertyAccessor == null)
            {
                return null;
            }

            const string smtpProperty = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
            var value = propertyAccessor.GetProperty(smtpProperty);
            return value as string;
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Errore lettura SMTP mailbox delegata: {ex.Message}");
            return null;
        }
    }

    
    
    
    private static string GetAccountTypeName(int accountType)
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
}
