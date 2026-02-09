# OnlyFirmaOutlook

OnlyFirmaOutlook è un'applicazione WPF per Windows che trasforma documenti Word in firme Outlook pronte all'uso. L'app guida l'utente passo-passo, gestisce i preset, crea backup automatici e consente il ripristino delle firme in caso di necessità.

## Funzionalità principali

- **Conversione Word → firme Outlook** con generazione di HTML/RTF/TXT e ricostruzione degli asset.
- **Preset**: selezione rapida dei modelli Word presenti nella cartella `media` o caricamento di file custom.
- **Modifica assistita in Word**: apertura del documento, verifica salvataggio e controllo stato.
- **Opzioni HTML** (filtrato o completo) per bilanciare compatibilità e fedeltà visiva.
- **Fix Outlook 2512+** per correggere i bordi delle tabelle nelle firme.
- **Gestione firme esistenti** con avvisi di sovrascrittura e cancellazione rapida.
- **Backup automatici** quando la destinazione è la cartella predefinita di Outlook, con ripristino e pulizia.
- **Log operativo** con copia/pulizia del file di log.

## Requisiti

- **Windows** (app WPF).
- **Microsoft Word** installato per l'editing e la conversione.
- **Microsoft Outlook Classic** per l'utilizzo delle firme generate.
- Per lo **sviluppo**: .NET SDK 8.0 e PowerShell.

## Percorsi predefiniti

- Cartella firme Outlook: `%APPDATA%\Microsoft\Signatures`.
- Output alternativo (quando Outlook non è disponibile o si sceglie un'altra destinazione): `%USERPROFILE%\Documents\OnlyFirmaOutlook\Output`.

## Flusso di lavoro dettagliato

### 1) Import del documento

- **Preset**: i preset sono letti dalla cartella `media` dell'app (distribuiti con la build). La selezione di un preset crea una copia temporanea locale e prepara l'editor.
- **File personalizzato**: sono accettati file `.doc`, `.docx` e `.rtf`. I file da rete vengono copiati in locale per evitare blocchi durante l'editing.
- **Normalizzazione nome firma**: il nome proposto viene ripulito da caratteri non validi per evitare problemi durante l'export.

### 2) Modifica in Word

- Il documento viene aperto in Word da una cartella temporanea dedicata.
- Lo stato dell'editing è monitorato; la conversione è abilitata solo dopo il salvataggio.

### 3) Export firme

- **HTML**: generato in formato filtrato o completo a seconda dell'opzione scelta.
- **RTF/TXT**: esportati per compatibilità con Outlook Classic.
- **Normalizzazione HTML**: rimozione di stili superflui, inline CSS, correzione bordi tabelle e ricostruzione degli asset.
- **Asset**: le immagini vengono incorporate o ricostruite nella cartella `<firma>_files`.
- **Backup**: se la destinazione è la cartella Outlook, viene creato un backup ZIP prima di sovrascrivere.

## Opzioni e filtri

- **HTML Filtrato**: riduce gli stili Microsoft/Word non necessari per migliorare la compatibilità.
- **HTML Completo**: preserva più stili di Word (utile quando serve maggiore fedeltà visiva).
- **Fix bordi tabelle Outlook 2512+**: aggiunge reset CSS/MSO a tabelle e celle per evitare bordi indesiderati introdotti da Outlook Classic 2512+.

## Uso rapido (utente finale)

1. Seleziona un preset oppure carica un documento Word.
2. Configura nome firma e account/identificativo.
3. Verifica la cartella di destinazione.
4. Apri in Word, modifica e salva.
5. Scegli formato HTML e opzioni di correzione.
6. Controlla eventuali firme esistenti, quindi converti e salva.
7. Verifica in Outlook che la firma sia corretta.

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

Lo script `scripts/build.ps1` gestisce pulizia, restore, build, test e publish per entrambi i runtime e copia i preset nella cartella di output.

```powershell
# Build+publish completo
./scripts/build.ps1 -Configuration Release

# Build senza test
./scripts/build.ps1 -Configuration Release -SkipTests

# Solo build (senza publish)
./scripts/build.ps1 -Configuration Release -SkipPublish

# Output personalizzato e senza bootstrapper
./scripts/build.ps1 -Configuration Release -OutputDir dist -SkipBootstrapper

# Skip copia preset
./scripts/build.ps1 -Configuration Release -SkipMediaCopy
```

### Script di pulizia

```powershell
# Pulizia standard (bin/obj + dist)
./scripts/clean.ps1

# Pulizia completa
./scripts/clean.ps1 -All

# Pulizia con rimozione dati utente (EditorTemp/Logs)
./scripts/clean.ps1 -IncludeUserData
```

## Distribuzione

- Pubblicare la cartella di output su share di rete.
- Gli utenti avviano **OnlyFirmaOutlook.Launcher.exe**, che rileva la bitness di Office e lancia la build corretta.

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

- **File Word su rete**: vengono copiati in locale per evitare blocchi durante la modifica.
- **Backup automatico**: creato solo se la destinazione è la cartella Outlook.
- **Outlook non installato**: scegliere una cartella alternativa e usare l'output manualmente.
- **Log e pulizia**: usa i pulsanti di log per copia e reset; i file temporanei vengono rimossi all'uscita.

## License

Progetto interno. Aggiungere qui la licenza se prevista.
