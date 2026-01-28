# OnlyFirmaOutlook

Applicazione Windows per convertire documenti Word in firme email Outlook.

## Funzionalità

- Conversione documenti Word (.doc/.docx) in firme Outlook (HTML, RTF, TXT)
- Gestione automatica immagini e asset incorporati
- Normalizzazione HTML per compatibilità Outlook (rimozione bordi tabelle, stili Word)
- Preset di firma preconfigurati
- Backup e ripristino firme esistenti
- Rilevamento account Outlook configurati (Exchange, IMAP, POP3, deleghe)
- Launcher automatico per Office 32/64-bit

## Requisiti

- Windows 10/11
- Microsoft Word (per conversione documenti)
- Microsoft Outlook (opzionale, per gestione firme)
- .NET 8 SDK (solo per sviluppo)

## Installazione

### Utente finale

1. Scaricare la cartella `dist` dalla release
2. Eseguire `OnlyFirmaOutlook.Launcher.exe`

Il launcher rileva automaticamente la versione di Office installata (32/64-bit) e avvia la build corretta.

### Sviluppatore

```powershell
git clone https://github.com/user/OnlyFirmaOutlook.git
cd OnlyFirmaOutlook
dotnet restore
dotnet run --project src/OnlyFirmaOutlook
```

## Struttura progetto

```
OnlyFirmaOutlook/
├── src/
│   ├── Bootstrapper/          # Launcher per selezione build 32/64-bit
│   ├── OnlyFirmaOutlook/      # Applicazione WPF principale
│   │   ├── Models/            # Classi dati
│   │   ├── Services/          # Logica business
│   │   ├── Views/             # Interfaccia utente
│   │   └── media/             # Preset firma predefiniti
│   └── Shared/                # Codice condiviso
├── tests/
│   └── OnlyFirmaOutlook.Tests/
├── scripts/
│   ├── build.ps1              # Script build e publish
│   └── clean.ps1              # Pulizia artefatti
└── dist/                      # Output build (generato)
```

## Comandi

### Build

```powershell
.\scripts\build.ps1
```

Opzioni:
- `-Configuration Debug|Release` (default: Release)
- `-SkipTests` salta esecuzione test
- `-SkipPublish` solo build senza publish
- `-Runtimes @("win-x86", "win-x64")` runtime target

### Test

```powershell
dotnet test
```

### Pulizia

```powershell
.\scripts\clean.ps1
```

## Configurazione

| Percorso | Descrizione |
|----------|-------------|
| `%APPDATA%\Microsoft\Signatures` | Cartella firme Outlook |
| `%LOCALAPPDATA%\OnlyFirmaOutlook\Logs` | Log applicazione |
| `<app>\media\` | Preset firma (.doc/.docx) |

## Distribuzione

1. Eseguire `.\scripts\build.ps1`
2. Copiare la cartella `dist\` su share di rete o PC target
3. Gli utenti eseguono `OnlyFirmaOutlook.Launcher.exe`

Contenuto `dist/`:
```
dist/
├── OnlyFirmaOutlook.Launcher.exe    # Launcher (eseguire questo)
├── win-x86/                         # Build per Office 32-bit
│   ├── OnlyFirmaOutlook.exe
│   └── media/
└── win-x64/                         # Build per Office 64-bit
    ├── OnlyFirmaOutlook.exe
    └── media/
```

## Preset firma

Per aggiungere preset predefiniti:

1. Creare documenti Word (.doc/.docx) con il layout della firma
2. Copiare in `src/OnlyFirmaOutlook/media/`
3. Ricompilare con `.\scripts\build.ps1`

I preset appaiono nella lista "Preset" dell'applicazione.

## Tecnologie

- .NET 8 / WPF
- HtmlAgilityPack (parsing HTML)
- PreMailer.Net (CSS inlining)
- Microsoft.Office.Interop.Word (conversione documenti)

## Licenza

MIT
