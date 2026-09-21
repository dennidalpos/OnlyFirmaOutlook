# Project Specification

## Goal
OnlyFirmaOutlook e un'app desktop WPF per Windows che converte documenti Word in firme Outlook pronte all'uso, con supporto a editing assistito, preset, backup e ripristino.

## Scope
- Import di documenti `.doc`, `.docx` e `.rtf` da preset locali o file scelti dall'utente.
- Apertura e modifica del documento in Microsoft Word tramite automazione Office.
- Esportazione della firma nei formati HTML, RTF e TXT compatibili con Outlook Classic.
- Normalizzazione dell'HTML e gestione di immagini locali nel formato nativo delle firme Outlook (`<firma>_files`).
- Gestione della cartella firme di Outlook o di una cartella di output alternativa.
- Backup ZIP delle firme esistenti e funzioni di ripristino snapshot.
- Script PowerShell standardizzati (`Build-App.ps1`, `Clean-App.ps1`, `Install-InstallerToolchain.ps1`, `Package-App.ps1`).
- Packaging canonico con Inno Setup (EXE) e WiX (MSI).
- Workflow CI GitHub Actions per restore, build e test su Windows.
- Test unitari sui servizi principali.

## Non Scope
- Supporto multipiattaforma diverso da Windows.
- Supporto a client email diversi da Outlook Classic.
- Editing interno del documento senza Microsoft Word installato.
- Sincronizzazione cloud, servizi web o componenti server-side.

## Architecture
- `src/OnlyFirmaOutlook`: applicazione WPF principale (`net10.0-windows`) con viste, modelli e servizi per conversione, installazione firme, immagini inline Outlook, logging e gestione file temporanei; `MainWindow` è suddivisa in partial class per separare editor, gestione firme e chrome UI.
- `src/Bootstrapper`: launcher che rileva la bitness di Office e avvia la build corretta.
- `src/Shared`: codice condiviso per il rilevamento della bitness di Office.
- `tests/OnlyFirmaOutlook.Tests`: progetto di test xUnit per repository e servizi.
- `packaging`: configurazioni di installer per Inno Setup (`packaging/innosetup/setup.iss`) e WiX (`packaging/msi/OnlyFirmaOutlook.wxs`).
- `scripts`: script PowerShell per build/publish (`Build-App.ps1`), pulizia (`Clean-App.ps1`), bootstrap toolchain (`Install-InstallerToolchain.ps1`) e packaging installer (`Package-App.ps1`).
- `.github/workflows/ci.yml`: pipeline CI Windows che verifica restore, build e test della soluzione.

## Constraints
- Richiede Windows, .NET 10 SDK per lo sviluppo e Microsoft Word installato per il workflow di conversione/editing.
- Il target applicativo e `net10.0-windows`; i runtime supportati in publish sono `win-x86` e `win-x64`.
- La soluzione usa `Microsoft.Office.Interop.Word`, quindi dipende da Office installato e disponibile localmente.
- I preset distribuiti sono letti dalla cartella `src/OnlyFirmaOutlook/media`.
- Le immagini delle firme usano riferimenti relativi alla cartella `<firma>_files`; l'app imposta per l'utente corrente l'opzione Outlook `Send Pictures With Document` per farle inviare come contenuti inline. Outlook va riavviato dopo l'applicazione dell'impostazione.
