# OnlyFirmaOutlook

OnlyFirmaOutlook è un'applicazione Windows che converte documenti Word in firme per Outlook Classic, richiedendo l'apertura in Microsoft Word per la modifica obbligatoria prima della conversione.

## Funzionalità principali

- Conversione di documenti Word (.doc/.docx) in firme Outlook (HTML/RTF/TXT)
- Apertura di Microsoft Word per la modifica obbligatoria prima della conversione
- Preset gestiti tramite cartella media e file personalizzati
- Rilevamento automatico account Outlook (opzionale)
- Supporto build x86/x64 con launcher che rileva la bitness di Office
- Esecuzione da share di rete con gestione file temporanei

## Requisiti

- Windows 10/11
- Microsoft Word 2013+ (2013, 2016, 2019, 2021, 365)
- .NET 8.0 Runtime (incluso nelle build self-contained)
- Microsoft Outlook (opzionale)

## Avvio rapido

1. Avvia `OnlyFirmaOutlook.Launcher.exe` dalla cartella di distribuzione.
2. Seleziona un preset o carica un documento Word personalizzato.
3. Modifica il documento in Word e salva.
4. Inserisci il nome della firma e completa la conversione.

## Preset

I preset sono documenti Word nella cartella `media` dell'applicazione.

- Percorso runtime: `AppContext.BaseDirectory\media`
- I file temporanei `~$` di Word vengono ignorati.

Per distribuirli in entrambe le build, inserisci i file in `src/OnlyFirmaOutlook/media` e usa lo script di build.

## Build

Requisiti sviluppo:

- Visual Studio 2022+
- .NET 8.0 SDK
- PowerShell 5.1+

Comandi principali:

```powershell
# Build e publish
.\scripts\build.ps1

# Build in Debug
.\scripts\build.ps1 -Configuration Debug

# Pulizia artefatti
.\scripts\clean.ps1
```

Output atteso:

```
dist/
├── OnlyFirmaOutlook.Launcher.exe
├── win-x86/
│   └── OnlyFirmaOutlook.exe
└── win-x64/
    └── OnlyFirmaOutlook.exe
```

## Struttura progetto

```
src/
├── OnlyFirmaOutlook/    # App WPF principale
└── Bootstrapper/        # Launcher per bitness
scripts/                 # Script di build/clean
```

## Log

I log sono salvati in:

```
%LOCALAPPDATA%\OnlyFirmaOutlook\Logs\app.log
```

## Licenza

MIT.
