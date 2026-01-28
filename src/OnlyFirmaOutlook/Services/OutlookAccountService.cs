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
                        var outlookAccount = new OutlookAccount
                        {
                            DisplayName = account.DisplayName ?? "Account senza nome",
                            SmtpAddress = account.SmtpAddress ?? string.Empty,
                            AccountType = GetAccountTypeName(account.AccountType)
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
            
            CleanupComObjects(accounts, session, outlookApp);
        }

        return result;
    }

    
    
    
    private void CleanupComObjects(dynamic? accounts, dynamic? session, dynamic? outlookApp)
    {
        _logger.Log("Cleanup oggetti COM Outlook...");

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
