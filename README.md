# OnlyFirmaOutlook

OnlyFirmaOutlook è un'applicazione WPF per Windows che trasforma documenti Word in firme Outlook pronte all'uso. L'app guida l'utente passo-passo, gestisce i preset, crea backup automatici e consente il ripristino delle firme in caso di necessità.

## Funzionalità principali

- **Conversione Word → firme Outlook** con generazione di HTML/RTF/TXT, asset compatibili e pubblicazione transazionale.
- **Preset**: selezione rapida dei modelli Word/RTF presenti nella cartella `media` o caricamento di file custom.
- **Modifica assistita in Word**: apertura del documento, verifica salvataggio e controllo stato.
- **Opzioni HTML** (filtrato o completo) per bilanciare compatibilità e fedeltà visiva.
- **Gestione firme esistenti** con avvisi di sovrascrittura e cancellazione rapida.
- **Backup obbligatori prima della sovrascrittura** nella cartella predefinita di Outlook, con ripristino snapshot e pulizia.
- **Log operativo** con copia/pulizia del file di log.

## Requisiti

- **Windows** (app WPF).
- **Microsoft Word** installato per l'editing e la conversione.
- **Microsoft Outlook Classic** per l'utilizzo delle firme generate.
- Per lo **sviluppo**: .NET SDK 10.0 e PowerShell.

## Percorsi predefiniti

- Cartella firme Outlook: `%APPDATA%\Microsoft\Signatures`.
- Output alternativo (quando Outlook non è disponibile o si sceglie un'altra destinazione): `%USERPROFILE%\Documents\OnlyFirmaOutlook\Output`.

## Flusso di lavoro dettagliato

### 1) Import del documento

- **Preset**: i preset sono letti dalla cartella `media` dell'app (distribuiti con la build). Sono supportati file `.doc`, `.docx` e `.rtf`. La selezione di un preset crea una copia temporanea locale e prepara l'editor.
- **File personalizzato**: sono accettati file `.doc`, `.docx` e `.rtf`. I file da rete vengono copiati in locale per evitare blocchi durante l'editing.
- **Normalizzazione nome firma**: il nome proposto viene ripulito da caratteri non validi per evitare problemi durante l'export.

### 2) Modifica in Word

- Il documento viene aperto in Word da una cartella temporanea dedicata.
- Lo stato dell'editing è monitorato; la conversione è abilitata solo dopo il salvataggio.

### 3) Export firme

- **HTML**: generato in formato filtrato o completo a seconda dell'opzione scelta.
- **RTF/TXT**: esportati per compatibilità con Outlook Classic.
- **Normalizzazione HTML**: rimozione di stili superflui e CSS inline, preservando il formato nativo delle immagini Word.
- **Immagini**: ogni immagine resta nella cartella `<firma>_files` con il nome generato da Word e un riferimento relativo nell'HTML. All'avvio l'app abilita l'impostazione Outlook che invia queste immagini come allegati inline (`cid:`).
- **Pubblicazione sicura**: conversione e normalizzazione avvengono in staging; gli artefatti esistenti sono sostituiti solo a conversione completata e vengono ripristinati se la pubblicazione fallisce.
- **Backup**: nella cartella Outlook, una sovrascrittura prosegue solo dopo la creazione riuscita del backup ZIP.
- **Ripristino backup**: il restore riallinea la cartella firme allo snapshot del backup, rimuovendo artefatti residui non presenti nell'archivio.

## Opzioni e filtri

- **HTML Filtrato**: riduce gli stili Microsoft/Word non necessari per migliorare la compatibilità.
- **HTML Completo**: preserva più stili di Word (utile quando serve maggiore fedeltà visiva).

## Uso rapido (utente finale)

1. Seleziona un preset oppure carica un documento Word.
2. Configura nome firma e account/identificativo.
3. Verifica la cartella di destinazione.
4. Apri in Word, modifica e salva.
5. Scegli il formato HTML.
6. Controlla eventuali firme esistenti, quindi converti e salva.
7. Chiudi e riapri Outlook Classic dopo il primo avvio dell'app, quindi verifica la firma e invia un messaggio di prova.

## Build e publish

### Build manuale (sviluppo)

```bash
# Build soluzione
$ dotnet build OnlyFirmaOutlook.sln -c Release

# Test
$ dotnet test OnlyFirmaOutlook.sln -c Release
```

### Continuous Integration

La pipeline GitHub Actions in `.github/workflows/ci.yml` esegue `restore`, `build`, `test` e una verifica del flusso `scripts/Build-App.ps1`, inclusi publish `win-x86`/`win-x64`, bootstrapper e copia preset.

### Publish manuale

```bash
# Publish app principale (x64 e x86)
$ dotnet publish src/OnlyFirmaOutlook/OnlyFirmaOutlook.csproj -c Release -r win-x64 --self-contained true
$ dotnet publish src/OnlyFirmaOutlook/OnlyFirmaOutlook.csproj -c Release -r win-x86 --self-contained true

# Publish launcher
$ dotnet publish src/Bootstrapper/Bootstrapper.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
```

### Script PowerShell (consigliato)

Gli script del repository seguono la convenzione `Verb-Noun.ps1` (con wrapper di retrocompatibilità `build.ps1` e `clean.ps1`):

#### 1) Bootstrap toolchain di packaging
Verifica e installa i prerequisiti (PowerShell 7, winget, .NET 10 SDK, Inno Setup 6, WiX Toolset):
```powershell
# Verifica stato dei prerequisiti
./scripts/Install-InstallerToolchain.ps1

# Installazione automatica prerequisiti mancanti via winget/dotnet tool
./scripts/Install-InstallerToolchain.ps1 -Install
```

#### 2) Build e publish applicazione
Lo script `scripts/Build-App.ps1` gestisce pulizia, restore, build, test e publish per entrambi i runtime (`win-x86` e `win-x64`) e copia i preset `.doc`, `.docx` e `.rtf` nella cartella di output:
```powershell
# Build+publish completo (con scelta modalità Framework-dependent / Self-contained)
./scripts/Build-App.ps1 -Configuration Release

# Build+publish forzando Framework-dependent
./scripts/Build-App.ps1 -Configuration Release -PublishMode FrameworkDependent

# Build senza rieseguire i test
./scripts/Build-App.ps1 -Configuration Release -SkipTests

# Solo compilazione (senza publish in dist/)
./scripts/Build-App.ps1 -Configuration Release -SkipPublish
```

#### 3) Packaging installer (Inno Setup EXE e WiX MSI)
Lo script `scripts/Package-App.ps1` compila i pacchetti di installazione pronti per la distribuzione:
```powershell
# Compila sia l'installer EXE (Inno Setup) che l'installer MSI (WiX)
./scripts/Package-App.ps1

# Solo installer Inno Setup EXE
./scripts/Package-App.ps1 -Format InnoSetup -Version "1.0.0"

# Solo installer WiX MSI (riutilizzando artefatti esistenti in dist/)
./scripts/Package-App.ps1 -Format WiX -SkipBuild
```
Gli installer vengono generati nella cartella `packaging/output/`.

#### 4) Pulizia artefatti
```powershell
# Pulizia standard (bin/obj + dist + packaging/output + TestResults)
./scripts/Clean-App.ps1

# Pulizia completa (inclusi .vs e packages)
./scripts/Clean-App.ps1 -All

# Pulizia con rimozione dati utente locali (EditorTemp/Logs)
./scripts/Clean-App.ps1 -IncludeUserData
```

## Distribuzione

1. **Installer stand-alone**: distribuire `OnlyFirmaOutlook-Setup-<versione>.exe` o `OnlyFirmaOutlook-<versione>.msi` (generati in `packaging/output/`).
2. **Share di rete**: copiare l'intera cartella `dist/` sulla share di rete; gli utenti avviano `OnlyFirmaOutlook.Launcher.exe`, che rileva la bitness di Office e lancia l'eseguibile corretto (`win-x86` o `win-x64`).

## Struttura del repository

```
src/                  # Sorgenti C# .NET 10 (App WPF, Bootstrapper, Shared)
tests/                # Test unitari xUnit
packaging/            # Configurazioni e sorgenti installer
  innosetup/          # Script Inno Setup per installer EXE (setup.iss)
  msi/                # Sorgenti WiX per installer MSI (OnlyFirmaOutlook.wxs)
  output/             # Pacchetti generati (ignorati da git)
scripts/              # Script PowerShell di automazione (Verb-Noun)
.github/              # Workflow CI GitHub Actions
dist/                 # Output publish locale (ignorato da git)
```

## Documentazione

- `PROJECT_SPEC.md`: specifica tecnica, architettura e vincoli del progetto.
- `PROJECT_STATUS.json`: tracker dei task, stato di avanzamento e verifiche.
- `AGENTS.md`: fatti non derivabili, quirk di progetto e comandi verificati.

## Note operative e troubleshooting

- **File Word su rete**: vengono copiati in locale per evitare blocchi durante la modifica.
- **Backup prima della sovrascrittura**: richiesto solo quando si sovrascrive nella cartella Outlook.
- **Immagini assenti al destinatario**: chiudi Outlook, avvia OnlyFirmaOutlook per applicare l'impostazione di invio inline, rigenera la firma e riapri Outlook. Verifica inoltre di usare una build Office aggiornata.
- **Outlook non installato**: scegliere una cartella alternativa e usare l'output manualmente.
- **Log e pulizia**: usa i pulsanti di log per copia e reset; i file temporanei vengono rimossi all'uscita.

## License

Software proprietario. Vedere il file `LICENSE`.

## Copyright

Copyright (c) 2026 Danny Perondi.
