# OnlyFirmaOutlook

OnlyFirmaOutlook è un'applicazione WPF per Windows che trasforma documenti Word in firme Outlook pronte all'uso. L'app guida l'utente passo-passo, gestisce i preset, crea backup automatici e consente il ripristino delle firme in caso di necessità.

## Funzionalità principali

- **Conversione Word → firme Outlook** con generazione di HTML/RTF/TXT e ricostruzione degli asset. 【F:src/OnlyFirmaOutlook/Views/GuideWindow.xaml†L46-L95】
- **Preset**: selezione rapida dei modelli Word presenti nella cartella `media` o caricamento di file custom. 【F:src/OnlyFirmaOutlook/Views/GuideWindow.xaml†L26-L37】
- **Modifica assistita in Word**: apertura del documento, verifica salvataggio e controllo stato. 【F:src/OnlyFirmaOutlook/Views/GuideWindow.xaml†L50-L61】
- **Opzioni HTML** (filtrato o completo) per bilanciare compatibilità e fedeltà visiva. 【F:src/OnlyFirmaOutlook/Views/GuideWindow.xaml†L62-L66】
- **Fix Outlook 2512+** per correggere i bordi delle tabelle nelle firme. 【F:src/OnlyFirmaOutlook/Views/MainWindow.xaml†L181-L186】
- **Gestione firme esistenti** con avvisi di sovrascrittura e cancellazione rapida. 【F:src/OnlyFirmaOutlook/Views/GuideWindow.xaml†L67-L71】
- **Backup automatici** quando la destinazione è la cartella predefinita di Outlook, con ripristino e pulizia. 【F:src/OnlyFirmaOutlook/Views/GuideWindow.xaml†L72-L77】
- **Log operativo** con copia/pulizia del file di log. 【F:src/OnlyFirmaOutlook/Views/GuideWindow.xaml†L92-L95】

## Requisiti

- **Windows** (app WPF).
- **Microsoft Word** installato per l'editing e la conversione.
- **Microsoft Outlook Classic** per l'utilizzo delle firme generate (la guida include i passaggi di verifica).【F:src/OnlyFirmaOutlook/Views/GuideWindow.xaml†L86-L90】
- Per lo **sviluppo**: .NET SDK 8.0 e PowerShell.

## Percorsi predefiniti

- Cartella firme Outlook: `%APPDATA%\Microsoft\Signatures`.【F:src/OnlyFirmaOutlook/Services/SignatureRepository.cs†L24-L31】
- Output alternativo (quando Outlook non è disponibile o si sceglie un'altra destinazione): `%USERPROFILE%\Documents\OnlyFirmaOutlook\Output`.【F:src/OnlyFirmaOutlook/Services/SignatureRepository.cs†L33-L40】

## Uso rapido (utente finale)

1. **Seleziona un preset** oppure carica un documento Word.
2. **Configura il nome firma** e l'account/identificativo.
3. **Verifica la cartella di destinazione** (di default quella di Outlook).
4. **Apri in Word**, modifica e salva il documento.
5. **Scegli il formato HTML** e le opzioni di correzione.
6. **Controlla eventuali firme esistenti**, quindi converti e salva.
7. **Verifica in Outlook** che la firma sia corretta.

La guida completa è disponibile nel pulsante “Guida” all'interno dell'app. 【F:src/OnlyFirmaOutlook/Views/GuideWindow.xaml†L20-L95】

## Build e publish

### Build manuale (sviluppo)

```bash
# Build soluzione
$ dotnet build OnlyFirmaOutlook.sln -c Release

# Test
$ dotnet test OnlyFirmaOutlook.sln -c Release
```

### Publish manuale

```bash
# Publish app principale (x64 e x86)
$ dotnet publish src/OnlyFirmaOutlook/OnlyFirmaOutlook.csproj -c Release -r win-x64 --self-contained true
$ dotnet publish src/OnlyFirmaOutlook/OnlyFirmaOutlook.csproj -c Release -r win-x86 --self-contained true

# Publish launcher
$ dotnet publish src/Bootstrapper/Bootstrapper.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
```

### Script PowerShell (consigliato)

Lo script `scripts/build.ps1` gestisce pulizia, restore, build, test e publish per entrambi i runtime e copia i preset nella cartella `dist`.

```powershell
# Build+publish completo
./scripts/build.ps1 -Configuration Release

# Build senza test
./scripts/build.ps1 -Configuration Release -SkipTests

# Solo build (senza publish)
./scripts/build.ps1 -Configuration Release -SkipPublish
```

Dettagli dello script: output in `dist`, launcher `OnlyFirmaOutlook.Launcher.exe`, cartelle `win-x86` e `win-x64`.【F:scripts/build.ps1†L1-L164】

## Distribuzione

- Pubblicare la cartella `dist` su share di rete.
- Gli utenti avviano **OnlyFirmaOutlook.Launcher.exe**, che rileva la bitness di Office e lancia la build corretta. 【F:src/Bootstrapper/Program.cs†L20-L88】

## Struttura del repository

```
src/
  Bootstrapper/       # Launcher che seleziona x86/x64 in base a Office
  OnlyFirmaOutlook/   # App WPF principale
  Shared/             # Componenti condivisi (rilevamento bitness Office)
tests/                # Test unitari
scripts/              # Script di build e pulizia
```

## Note operative e troubleshooting

- **File Word su rete:** i documenti provenienti da share vengono copiati in locale per evitare blocchi durante la modifica. 【F:src/OnlyFirmaOutlook/Views/GuideWindow.xaml†L26-L37】
- **Backup automatico:** creato solo se la destinazione è la cartella Outlook. 【F:src/OnlyFirmaOutlook/Views/GuideWindow.xaml†L72-L77】
- **Outlook non installato:** scegli una cartella alternativa e usa l'output come firma manuale.
- **Log e pulizia:** usa i pulsanti di log per copia e reset; i file temporanei vengono rimossi all'uscita. 【F:src/OnlyFirmaOutlook/Views/GuideWindow.xaml†L92-L95】

## License

Progetto interno. Aggiungere qui la licenza se prevista.
