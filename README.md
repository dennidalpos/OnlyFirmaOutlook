# OnlyFirmaOutlook

Utility Windows per convertire documenti Word in firme per Microsoft Outlook Classic (desktop).

## Funzionalità

- **Editor Word integrato** per la modifica obbligatoria delle firme prima della conversione
- Toolbar personalizzata con controlli completi di formattazione (font, grassetto, colori, allineamento, link, immagini)
- Converte documenti Word (.doc, .docx) in firme Outlook (HTML, RTF, TXT)
- Supporta preset predefiniti e documenti personalizzati
- I preset richiedono obbligatoriamente modifica nell'editor (volutamente incompleti)
- Rileva automaticamente gli account Outlook configurati
- Funziona anche senza Outlook installato (esportazione in cartella alternativa)
- Supporta Office 32-bit e 64-bit con build dedicate
- Eseguibile da share di rete UNC
- Interfaccia grafica WPF moderna

## Prerequisiti

### Obbligatori
- **Windows 10/11**
- **Microsoft Word** (qualsiasi versione: 2013, 2016, 2019, 2021, 365)
- **.NET 8.0 Runtime** (incluso nelle build self-contained)

### Opzionali
- **Microsoft Outlook** (per il rilevamento automatico degli account e la cartella firme predefinita)

### Note sulla bitness di Office
L'applicazione fornisce due build separate:
- `win-x86`: per Office 32-bit
- `win-x64`: per Office 64-bit

Il launcher (`OnlyFirmaOutlook.Launcher.exe`) rileva automaticamente la bitness di Office installato e avvia la versione corretta.

## Installazione

### Da share di rete (consigliato per ambienti aziendali)

1. Copiare l'intera cartella `dist\` su una share di rete accessibile:
   ```
   \\server\share\OnlyFirmaOutlook\
   ```

2. Gli utenti eseguono `OnlyFirmaOutlook.Launcher.exe` dalla share.

### Installazione locale

1. Copiare la cartella `dist\` sul PC locale
2. Eseguire `OnlyFirmaOutlook.Launcher.exe`

## Build da sorgente

### Prerequisiti per lo sviluppo
- Visual Studio 2022 o successivo
- .NET 8.0 SDK
- PowerShell 5.1 o successivo

### Comandi di build

```powershell
# Build e publish
.\scripts\build.ps1

# Build in configurazione Debug
.\scripts\build.ps1 -Configuration Debug

# Pulizia artefatti
.\scripts\clean.ps1

# Pulizia completa (inclusi .vs e file nascosti)
.\scripts\clean.ps1 -All
```

### Output del build

Dopo il build, la cartella `dist\` conterrà:
```
dist\
├── OnlyFirmaOutlook.Launcher.exe    # Launcher (avvia la build corretta)
├── win-x86\                          # Build 32-bit
│   ├── OnlyFirmaOutlook.exe
│   ├── media\                        # Cartella per i preset
│   └── ...
└── win-x64\                          # Build 64-bit
    ├── OnlyFirmaOutlook.exe
    ├── media\                        # Cartella per i preset
    └── ...
```

## Utilizzo

### Flusso completo (Con Outlook installato)

1. **Avviare** `OnlyFirmaOutlook.Launcher.exe`
2. **Selezionare** un preset dalla lista o caricare un documento Word personalizzato
   - L'editor Word si aprirà automaticamente (modale)
3. **Modificare** il documento nell'editor integrato:
   - Utilizzare la toolbar per formattare il testo (font, grassetto, colori, allineamento)
   - Inserire link e immagini
   - Personalizzare i campi segnaposto (nome, ruolo, contatti, ecc.)
4. **Salvare** il documento nell'editor cliccando "Salva e chiudi"
5. **Inserire** il nome della firma nella finestra principale
6. **Selezionare** l'account Outlook (opzionale, aggiunge l'email al nome della firma)
7. **Scegliere** il tipo di HTML (Filtrato consigliato)
8. **Cliccare** "Converti e salva"

La firma verrà salvata in `%APPDATA%\Microsoft\Signatures\` e sarà immediatamente disponibile in Outlook.

#### Note importanti sull'editor

- **Modifica obbligatoria**: Tutte le firme (preset e file personalizzati) devono passare dall'editor prima della conversione
- **Preset incompleti**: I preset sono volutamente incompleti e contengono segnaposto da personalizzare
- **Stato documento**: Il pulsante "Converti e salva" è disabilitato finché il documento non è stato aperto e salvato almeno una volta
- **Rimodifica**: È possibile riaprire l'editor cliccando su "Modifica firma" prima della conversione finale

### Senza Outlook installato

1. Avviare l'applicazione
2. Selezionare o caricare un documento Word
3. Inserire il nome della firma
4. Opzionalmente, inserire un identificativo (es. email)
5. La cartella di destinazione predefinita sarà:
   `%USERPROFILE%\Documents\OnlyFirmaOutlook\Output\`
6. Cliccare "Converti e salva"

#### Installazione manuale della firma

Per utilizzare la firma in Outlook su un altro PC:

1. Copiare i file generati:
   - `NomeFirma.htm`
   - `NomeFirma.rtf`
   - `NomeFirma.txt`
   - `NomeFirma_files\` (se presente)

2. Incollare in:
   ```
   %APPDATA%\Microsoft\Signatures\
   ```

3. Riavviare Outlook

### Preset (Media)

I preset sono documenti Word predefiniti che appaiono nella lista di selezione dell'applicazione.

#### Dove salvare i preset

I file `.doc` o `.docx` devono essere copiati nelle cartelle `media\` di **entrambe** le build:

```
dist\
├── win-x86\
│   └── media\                    # ← Copia qui i preset
│       ├── FirmaAziendale.docx
│       └── FirmaMinimale.docx
└── win-x64\
    └── media\                    # ← E anche qui (stessi file)
        ├── FirmaAziendale.docx
        └── FirmaMinimale.docx
```

#### Metodi per aggiungere preset

**Metodo 1 - Copia manuale (dopo il build):**
1. Crea il documento Word con la firma desiderata
2. Copia il file in `dist\win-x86\media\`
3. Copia lo stesso file in `dist\win-x64\media\`
4. Riavvia l'applicazione

**Metodo 2 - Copia automatica (durante il build):**
1. Crea il documento Word con la firma desiderata
2. Copia il file in `src\OnlyFirmaOutlook\media\`
3. Esegui `.\scripts\build.ps1`
4. I file verranno copiati automaticamente in entrambe le build

#### Note sui preset

- L'applicazione legge i preset da `AppContext.BaseDirectory\media\`
- I file che iniziano con `~$` (temporanei di Word) vengono ignorati
- Il nome visualizzato è il nome del file senza estensione

## Editor Word Integrato

### Architettura tecnica

#### Scelta tecnologica: Word Embedded
L'editor utilizza **Microsoft.Office.Interop.Word** con Word embedded nella finestra WPF tramite `WindowsFormsHost`.

**Motivazione**:
- **Controllo completo**: Gestione diretta del ciclo di vita dell'istanza Word
- **Stabilità**: Nessuna interferenza con altre istanze Word dell'utente
- **Integrazione UI**: Word integrato nell'interfaccia dell'applicazione
- **Thread STA dedicato**: Tutte le operazioni COM su thread separato

#### Componenti

**WordEditorWindow** (Views/):
- Finestra modale WPF per editing
- Host Word tramite `WindowsFormsHost` e embedding HWND
- Gestione eventi salvataggio e chiusura
- Thread STA dedicato per Word COM

**WordToolbar** (Controls/):
- Toolbar custom WPF (non Ribbon Office)
- Controlli diretti via COM Word:
  - **Font**: Dropdown con tutti i font di sistema, dimensione editabile
  - **Formattazione**: Grassetto, corsivo, sottolineato, colore testo
  - **Paragrafo**: Allineamento (sx, centro, dx)
  - **Inserimenti**: Link (con dialog), immagini (file picker)
  - **Editing**: Undo/Redo
  - **Zoom**: Aumenta/Riduci (10% step, range 10-500%)
- Nessuna funzionalità non supportata dal COM Word classico

**WordEditorService** (Services/):
- Prepara file per editing (copia in `EditorTemp\{guid}\`)
- Cleanup cartelle temporanee post-conversione
- Validazione stato editor

**EditorState** (Models/):
- Stato modifica documento:
  - `IsDocumentOpened`: Documento aperto almeno una volta
  - `IsDocumentSaved`: Documento salvato almeno una volta
  - `IsReadyForConversion`: Pronto per conversione (aperto + salvato)
  - `HasUnsavedChanges`: Modifiche non salvate

#### Flusso operativo

1. **Selezione firma** → `WordEditorService.PrepareFileForEditing()`
   - Crea `EditorTemp\{guid}\`
   - Copia file locale
   - Crea `EditorState`

2. **Apertura WordEditorWindow** (modale)
   - Thread STA dedicato per Word COM
   - Word embedded via HWND
   - Toolbar collegata al documento

3. **Modifica e salvataggio**
   - Toolbar → Comandi COM → Word
   - Evento `DocumentModified` → Aggiorna stato
   - Pulsante "Salva" → `Document.Save()` → Marca `IsDocumentSaved`

4. **Chiusura editor**
   - Se modifiche non salvate → Dialog conferma
   - Cleanup COM: `Document.Close()` → `Application.Quit()` → `Marshal.FinalReleaseComObject`
   - Forza GC per rilascio definitivo

5. **Conversione** (solo se `IsReadyForConversion`)
   - `WordConversionService.ConvertDocument(editorState.LocalFilePath)`
   - File sorgente: copia modificata in EditorTemp

6. **Cleanup post-conversione**
   - `WordEditorService.CleanupEditorTempFolder(sessionId)`
   - Eliminazione cartella `EditorTemp\{guid}\` e contenuto

#### Threading

- **Thread principale (STA)**: UI WPF
- **Thread editor (STA)**: Word COM operations
- **Sincronizzazione**: `ManualResetEvent` per load, `Dispatcher.Invoke` per UI updates

#### Caricamento font di sistema

```csharp
var fonts = new System.Drawing.Text.InstalledFontCollection();
var fontNames = fonts.Families.Select(f => f.Name).OrderBy(n => n);
```

**Font default**: Calibri → Arial → primo disponibile

### Limitazioni note

- **Nessun supporto Ribbon Office**: Solo comandi COM diretti
- **Formattazione avanzata limitata**: Tabelle complesse, stili Word avanzati richiedono Word completo
- **Single document**: Un documento alla volta nell'editor
- **No collaborative editing**: Editing locale, nessuna sincronizzazione cloud

## Troubleshooting

### Errore COM Interop

**Sintomo**: Errore durante la conversione con codice 0x800A...

**Cause possibili**:
- Office non installato correttamente
- Bitness dell'applicazione non corrisponde a Office
- Word in esecuzione con documenti aperti

**Soluzioni**:
1. Verificare che Word sia installato e funzionante
2. Chiudere tutti i documenti Word aperti
3. Utilizzare il launcher per avviare la build corretta

### Office bitness mismatch

**Sintomo**: Errore "Impossibile creare istanza di Word" o errore COM generico

**Causa**: L'applicazione x86 non può comunicare con Office x64 e viceversa

**Soluzione**: Utilizzare sempre `OnlyFirmaOutlook.Launcher.exe` che rileva automaticamente la bitness corretta

### Permessi cartella firme

**Sintomo**: "Cartella non scrivibile"

**Causa**: Permessi insufficienti sulla cartella `%APPDATA%\Microsoft\Signatures\`

**Soluzioni**:
1. Verificare di avere permessi di scrittura sulla cartella
2. Selezionare una cartella alternativa con il pulsante "Sfoglia..."
3. Contattare l'amministratore di sistema

### Share UNC e file lock

**Sintomo**: Errore durante la selezione di un preset da share di rete

**Causa**: Il file potrebbe essere bloccato da un altro processo

**Soluzioni**:
1. L'applicazione copia automaticamente i file in locale prima di elaborarli
2. Verificare che il file non sia aperto da altri utenti
3. Verificare la connettività alla share di rete

### Protected View

**Sintomo**: Errore 0x800A175D durante la conversione

**Causa**: Il documento è in "Protected View" (Visualizzazione protetta) di Word

**Soluzioni**:
1. Aprire il documento manualmente in Word
2. Cliccare "Abilita modifica" nella barra gialla
3. Salvare e chiudere il documento
4. Riprovare la conversione

In alternativa, disabilitare Protected View per i file dalla share di rete:
1. Aprire Word > File > Opzioni > Centro protezione
2. Impostazioni Centro protezione > Visualizzazione protetta
3. Deselezionare le opzioni relative ai percorsi di rete

### Word zombie process

**Sintomo**: Processi WINWORD.EXE rimangono attivi dopo la conversione

**Causa**: Errore durante il cleanup degli oggetti COM

**Soluzioni**:
1. L'applicazione implementa un cleanup rigoroso con `Marshal.FinalReleaseComObject`
2. In caso di crash, terminare manualmente i processi Word dal Task Manager
3. Verificare che non ci siano altri software che interferiscono con Word

### Editor non si apre / Word non risponde

**Sintomo**: L'editor si blocca durante il caricamento o Word non appare

**Cause possibili**:
- Word già aperto con documenti bloccati
- Istanza Word zombie da sessione precedente
- Conflitto con add-in Word

**Soluzioni**:
1. Chiudere tutte le istanze di Word dal Task Manager
2. Riavviare l'applicazione
3. Verificare che Word si apra normalmente in modalità standalone
4. Disabilitare temporaneamente add-in Word di terze parti

### Toolbar non risponde ai comandi

**Sintomo**: I pulsanti della toolbar non hanno effetto sul documento

**Causa**: Perdita del riferimento COM al documento Word

**Soluzioni**:
1. Chiudere e riaprire l'editor
2. Se persiste, riavviare l'applicazione
3. Verificare che Word non sia in "Protected View"

### Documento non modificabile nell'editor

**Sintomo**: Il documento appare read-only nell'editor

**Causa**: Attributi file o permessi insufficienti

**Soluzioni**:
1. L'applicazione rimuove automaticamente l'attributo read-only
2. Verificare permessi sulla cartella `%LOCALAPPDATA%\OnlyFirmaOutlook\EditorTemp\`
3. Eseguire l'applicazione con permessi sufficienti

### Conversione bloccata dopo editing

**Sintomo**: Il pulsante "Converti e salva" rimane disabilitato dopo aver modificato

**Causa**: Documento non salvato nell'editor

**Soluzioni**:
1. Riaprire l'editor con "Modifica firma"
2. Cliccare "Salva" o "Salva e chiudi" nell'editor
3. Lo stato deve mostrare "Modificata e pronta"

### Cartelle EditorTemp non eliminate

**Sintomo**: Cartelle `%LOCALAPPDATA%\OnlyFirmaOutlook\EditorTemp\{guid}\` non vengono eliminate

**Causa**: File bloccati da Word o cleanup fallito

**Soluzioni**:
1. Le cartelle più vecchie di 1 giorno vengono eliminate automaticamente all'avvio
2. Chiudere l'applicazione e tutte le istanze Word
3. Eliminare manualmente le cartelle se necessario
4. Riavviare l'applicazione

### Log e diagnostica

I log dell'applicazione sono salvati in:
```
%LOCALAPPDATA%\OnlyFirmaOutlook\Logs\app.log
```

Per visualizzare il log:
1. Utilizzare il pulsante "Apri file log" nell'interfaccia
2. Oppure navigare manualmente alla cartella

Eventi loggati per l'editor:
- Apertura editor
- Copia file locale in EditorTemp
- Salvataggi
- Chiusura editor
- Cleanup cartelle temporanee
- Errori COM Word

## Struttura del progetto

```
OnlyFirmaOutlook/
├── src/
│   ├── OnlyFirmaOutlook/           # Applicazione principale WPF
│   │   ├── Models/                  # Modelli dati
│   │   │   ├── EditorState.cs       # Stato editor Word
│   │   │   ├── SignatureInfo.cs
│   │   │   ├── OutlookAccount.cs
│   │   │   └── PresetFile.cs
│   │   ├── Services/                # Servizi business logic
│   │   │   ├── WordEditorService.cs # Gestione editor integrato
│   │   │   ├── WordConversionService.cs
│   │   │   ├── LoggingService.cs
│   │   │   ├── TempFileManager.cs
│   │   │   ├── PresetService.cs
│   │   │   ├── OutlookAccountService.cs
│   │   │   ├── SignatureRepository.cs
│   │   │   └── OfficeBitnessDetector.cs
│   │   ├── Windows/                 # Finestre WPF
│   │   │   └── WordEditorWindow.xaml/cs  # Editor Word modale
│   │   ├── Views/                   # Views principale
│   │   │   └── MainWindow.xaml/cs
│   │   ├── ViewModels/              # ViewModels
│   │   │   └── WordEditorViewModel.cs
│   │   ├── Controls/                # Controlli custom
│   │   │   └── WordToolbar.xaml/cs  # Toolbar formattazione Word
│   │   ├── Helpers/                 # Utility
│   │   │   └── ComHelper.cs         # Helper pulizia COM objects
│   │   ├── Styles/                  # Stili XAML
│   │   └── media/                   # Preset predefiniti
│   └── Bootstrapper/               # Launcher per rilevamento bitness
├── scripts/
│   ├── build.ps1                   # Script di build
│   └── clean.ps1                   # Script di pulizia
├── dist/                           # Output delle build (generato)
└── README.md
```

## Architettura

### Servizi principali

- **LoggingService**: Logging centralizzato su file e UI
- **TempFileManager**: Gestione file temporanei per esecuzione da share
- **WordEditorService**: Gestione editor Word integrato e cartelle temporanee dedicate
- **OfficeBitnessDetector**: Rilevamento bitness di Office
- **PresetService**: Caricamento preset dalla cartella media
- **OutlookAccountService**: Lettura account Outlook via COM
- **SignatureRepository**: Gestione firme esistenti
- **WordConversionService**: Conversione documenti Word

### Finestre e controlli

- **MainWindow**: Finestra principale dell'applicazione
- **WordEditorWindow**: Finestra modale per l'editing dei documenti Word
- **WordToolbar**: Toolbar personalizzata con controlli di formattazione via COM Word
- **WordEditorViewModel**: ViewModel per gestione stato editor

### Threading

- L'applicazione WPF esegue su thread STA (Single Thread Apartment)
- Le chiamate COM a Word e Outlook avvengono sul thread UI
- Le operazioni lunghe usano `Task.Run` ma ritornano sul thread UI per le chiamate COM

### Gestione share di rete e file temporanei

I file dalla share non vengono mai modificati direttamente:

**Flusso completo**:
1. Il preset/file viene copiato in `%LOCALAPPDATA%\OnlyFirmaOutlook\Temp\{guid}\` (TempFileManager)
2. Una seconda copia viene creata in `%LOCALAPPDATA%\OnlyFirmaOutlook\EditorTemp\{guid}\` (WordEditorService)
3. L'editor Word lavora **esclusivamente** sulla copia in EditorTemp
4. Dopo conversione riuscita, la cartella EditorTemp viene eliminata
5. Alla chiusura app, cleanup best-effort di tutte le cartelle temporanee
6. All'avvio, cleanup delle cartelle orfane (> 1 giorno)

**Motivazione doppia copia**: Separazione tra file di staging (Temp) e file di editing (EditorTemp) per garantire che le modifiche non influenzino i file sorgente.

## Licenza

Questo progetto è rilasciato con licenza MIT.

## Changelog

### 1.0.0
- Prima versione pubblica
- Supporto Office 32-bit e 64-bit
- Esecuzione da share di rete
- Interfaccia WPF
- Build self-contained
