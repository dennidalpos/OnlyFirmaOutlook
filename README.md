# OnlyFirmaOutlook

OnlyFirmaOutlook è un'applicazione desktop Windows (WPF) che converte documenti Word in firme Outlook (HTML/RTF/TXT), con gestione degli asset e installazione nella cartella firme. Il launcher seleziona automaticamente la build corretta in base alla bitness di Office installata.

## Requisiti

- Windows 10/11.
- Microsoft Word installato (necessario per la conversione).
- Microsoft Outlook opzionale (per rilevare account e gestire firme nella cartella predefinita).
- .NET 8 SDK per sviluppo, oppure build self-contained generate dallo script di publish.

## Configurazione

- Variabili d'ambiente: nessuna.
- Cartella firme Outlook predefinita: `%APPDATA%\Microsoft\Signatures`.
- Log applicativi: `%LOCALAPPDATA%\OnlyFirmaOutlook\Logs\app.log`.
- Preset Word runtime: `<cartella app>\media` (file `.doc`/`.docx`).
- Backup firme: file ZIP con prefisso `backup_firme_onlyfirmaoutlook_` nella cartella firme di Outlook.

## Struttura cartelle

```
OnlyFirmaOutlook.sln
README.md
scripts/
  build.ps1
  clean.ps1
src/
  Bootstrapper/
  OnlyFirmaOutlook/
  Shared/
tests/
  OnlyFirmaOutlook.Tests/
```

- `src/Bootstrapper`: launcher che seleziona `win-x86`/`win-x64` in base a Office.
- `src/OnlyFirmaOutlook`: applicazione WPF principale.
- `src/Shared`: componenti condivisi.

## Setup e sviluppo locale

1. Installare .NET 8 SDK.
2. Ripristinare le dipendenze:
   ```powershell
   dotnet restore .\OnlyFirmaOutlook.sln
   ```
3. Avviare l'app in debug:
   ```powershell
   dotnet run --project .\src\OnlyFirmaOutlook\OnlyFirmaOutlook.csproj
   ```

## Comandi principali

### Build
```powershell
.\scripts\build.ps1
```

### Test
```powershell
dotnet test .\OnlyFirmaOutlook.sln
```

### Lint
Non è presente un comando di lint nel repository.

### Dev
```powershell
dotnet run --project .\src\OnlyFirmaOutlook\OnlyFirmaOutlook.csproj
```

### Pulizia
```powershell
.\scripts\clean.ps1
```

## Esecuzione in produzione

1. Generare le build con `scripts/build.ps1` (output in `dist/`).
2. Distribuire l'intera cartella `dist` su PC locali o share di rete.
3. Avviare `OnlyFirmaOutlook.Launcher.exe`, che seleziona automaticamente la build corretta (`win-x86`/`win-x64`).

## Troubleshooting essenziale

- Word non installato o non accessibile: la conversione non può partire. Verificare installazione di Word e riprovare.
- Nessun preset visibile: creare la cartella `media` accanto all'eseguibile e aggiungere file `.doc/.docx`.
- Backup non creati: vengono generati solo nella cartella firme predefinita di Outlook.
