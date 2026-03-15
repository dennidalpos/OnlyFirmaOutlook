# Project Specification

## Goal
OnlyFirmaOutlook e un'app desktop WPF per Windows che converte documenti Word in firme Outlook pronte all'uso, con supporto a editing assistito, preset, backup e ripristino.

## Scope
- Import di documenti `.doc`, `.docx` e `.rtf` da preset locali o file scelti dall'utente.
- Apertura e modifica del documento in Microsoft Word tramite automazione Office.
- Esportazione della firma nei formati HTML, RTF e TXT compatibili con Outlook Classic.
- Normalizzazione dell'HTML e ricostruzione degli asset.
- Gestione della cartella firme di Outlook o di una cartella di output alternativa.
- Backup ZIP delle firme esistenti e funzioni di ripristino snapshot.
- Script PowerShell per build, test, clean e publish.
- Workflow CI GitHub Actions per restore, build e test su Windows.
- Test unitari sui servizi principali.

## Non Scope
- Supporto multipiattaforma diverso da Windows.
- Supporto a client email diversi da Outlook Classic.
- Editing interno del documento senza Microsoft Word installato.
- Sincronizzazione cloud, servizi web o componenti server-side.

## Architecture
- `src/OnlyFirmaOutlook`: applicazione WPF principale (`net8.0-windows`) con viste, view model, modelli e servizi per conversione, installazione firme, logging e gestione file temporanei; `MainWindow` e suddivisa in partial class per separare editor, gestione firme e chrome UI.
- `src/Bootstrapper`: launcher che rileva la bitness di Office e avvia la build corretta.
- `src/Shared`: codice condiviso per il rilevamento della bitness di Office.
- `tests/OnlyFirmaOutlook.Tests`: progetto di test xUnit per repository e servizi.
- `scripts`: script PowerShell per build/publish e pulizia del repository.
- `.github/workflows/ci.yml`: pipeline CI Windows che verifica restore, build e test della soluzione.

## Constraints
- Richiede Windows, .NET 8 SDK per lo sviluppo e Microsoft Word installato per il workflow di conversione/editing.
- Il target applicativo e `net8.0-windows`; i runtime supportati in publish sono `win-x86` e `win-x64`.
- La soluzione usa `Microsoft.Office.Interop.Word`, quindi dipende da Office installato e disponibile localmente.
- I preset distribuiti sono letti dalla cartella `src/OnlyFirmaOutlook/media`.
