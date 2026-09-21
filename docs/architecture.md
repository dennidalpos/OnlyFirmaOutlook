# Architettura Tecnica

Documentazione dell'architettura software, della pipeline di conversione e dei vincoli tecnici di OnlyFirmaOutlook (.NET 10 WPF).

---

## Mappa dei Progetti e Componenti

```
src/
├── OnlyFirmaOutlook/        # Applicazione WPF principale (net10.0-windows)
│   ├── Models/              # DTO e stati (EditorState, SignatureInfo, BackupInfo, etc.)
│   ├── Services/            # Logica di conversione, normalizzazione, repository e filesystem
│   ├── Styles/              # Temi e stili XAML
│   └── Views/               # MainWindow (divisa in partial class) e GuideWindow
├── Bootstrapper/            # Launcher console/WPF compatto per selezione automatica architettura
└── Shared/                  # Codice condiviso (OfficeBitnessDetector)

packaging/                   # Sorgenti di distribuzione installer
├── innosetup/               # Setup Inno Setup EXE (setup.iss)
└── msi/                     # Pacchetto WiX Toolset MSI (OnlyFirmaOutlook.wxs)
```

### Struttura di MainWindow (Partial Classes)
- **`MainWindow.xaml.cs`**: ciclo di vita finestra, rilevamento configurazione Office/Word, caricamento account e logica di avanzamento step.
- **`MainWindow.Editor.cs`**: gestione file temporanei di editing, `FileSystemWatcher` e timer di polling lock per monitorare salvataggio e chiusura di Word.
- **`MainWindow.SignatureManagement.cs`**: lista firme, eliminazione, calcolo nomi finali, backup ZIP e ripristino.
- **`MainWindow.Chrome.cs`**: gestione overlay di caricamento (busy), pannello log e finestra guida integrata.

---

## Rilevamento Bitness Office (Bootstrapper)

All'avvio della distribuzione unificata, l'utente esegue `OnlyFirmaOutlook.Launcher.exe`. Il launcher invoca `OfficeBitnessDetector`, che analizza il sistema attraverso 4 livelli gerarchici:

1. **Chiave Registro Outlook**: controlla `Software\Microsoft\Office\<v>\Outlook` (valore `Bitness` per versioni 16.0, 15.0, 14.0).
2. **Click-to-Run**: controlla `Software\Microsoft\Office\ClickToRun\Configuration` (valore `Platform`).
3. **Installazione MSI**: verifica il percorso di installazione di Office (`InstallRoot`) nel registro a 32 e 64 bit.
4. **PE Header Eseguibile Word**: ispeziona l'header dell'eseguibile `WINWORD.EXE` leggendo `IMAGE_FILE_HEADER.Machine` (`IMAGE_FILE_MACHINE_I386` vs `IMAGE_FILE_MACHINE_AMD64`).
5. **Fallback**: in assenza di riscontro univoco, seleziona `win-x64` come default sicuro.

Il bootstrapper avvia quindi il processo figlio corretto da `win-x86\OnlyFirmaOutlook.exe` o `win-x64\OnlyFirmaOutlook.exe`, inoltrando argomenti da riga di comando.

---

## Pipeline di Conversione Transazionale

La generazione della firma segue una sequenza a passaggi isolati per garantire che una conversione fallita non corrompa mai le firme esistenti:

```
[Documento Word]
       │
       ▼
 1. WordEditorService        Copia isolata in %LOCALAPPDATA%\OnlyFirmaOutlook\EditorTemp\<guid>
       │
       ▼
 2. Word.Application (COM)   Apertura headless e salvataggio in Staging (%LOCALAPPDATA%\...\Temp\<guid>)
       │                     Genera: .htm (filtrato o standard), .rtf, .txt e cartella asset
       ▼
 3. Normalizzazione HTML     WordHtmlSignatureNormalizer (rimozione tag e stili mso-* non visibili)
       │
       ▼
 4. CSS Inlining             CssInliner (inlining delle classi CSS da blocchi <style> negli elementi HTML)
       │
       ▼
 5. Gestione Asset           AssetManager (validazione anti path-traversal, copia immagini in <firma>_files/)
       │
       ▼
 6. Atomic Publish           Sostituzione atomica nella cartella di destinazione; rollback in caso di errore
```

### Componenti Chiave della Pipeline

- **`WordHtmlSignatureNormalizer`**:
  - Elimina nodi non renderizzabili o insicuri: `<script>`, `<meta>`, `<xml>`, `<o:p>`, commenti HTML e tag con namespace Word (`w:*`).
  - Pulisce gli attributi `style="..."`: rimuove proprietà proprietarie `mso-*` e `tab-stops`, preservando esplicitamente `mso-line-height-rule` (critico per l'interlinea in Outlook) e tutti gli attributi visivi standard (font, dimensioni, colori, margini, allineamenti).
- **`CssInliner`**: converte le regole CSS definite nei tag `<style>` in stili inline su ciascun elemento, garantendo compatibilità con i client email che non supportano CSS embedded.
- **`AssetManager`**:
  - Analizza tutti i nodi `<img>`.
  - Impedisce vulnerabilità di path-traversal: verifica che i percorsi risolti rimangano confinati nella directory base.
  - Copia le immagini locali nella sottocartella canonica `<firma>_files/` con riferimenti HTML relativi (`<img src="NomeFirma_files/immagine.png">`).
- **`TempCleanupHelper`**: incapsula la rimozione delle cartelle temporanee con tentativi ripetuti a ritardo progressivo per gestire i lock transitori rilasciati con ritardo da Word o da antivirus.

---

## Integrazione Registro Outlook

All'avvio, `OutlookInlineImageSettings` scrive nel ramo `HKCU\Software\Microsoft\Office\<v>\Outlook\Options\Mail` il valore DWORD:
```
"Send Pictures With Document" = 1
```
per le versioni Office 16.0, 15.0 e 14.0. Questo fa sì che Outlook Classic converta automaticamente i riferimenti relativi `<firma>_files` in allegati inline `cid:` all'invio del messaggio.

---

## Vincoli di Progetto e Non-Scope

- **Ambiente di esecuzione**: solo Windows (dipendenza da WPF, Win32 Registry, COM interop Word).
- **Client supportato**: esclusivamente Outlook Classic. Nessun supporto per "New Outlook" (client web-based Monarch), webmail o client terzi.
- **Dipendenze esterne**: richiede Microsoft Word installato localmente per convertire i documenti. Nessuna dipendenza da servizi cloud o server esterni.
