# OnlyFirmaOutlook

OnlyFirmaOutlook è un'applicazione desktop WPF (.NET 10) per Windows che converte documenti Microsoft Word (.docx, .doc, .rtf) in firme Outlook Classic pronte all'uso, con normalizzazione HTML, gestione degli asset grafici locali, backup automatici e packaging di installazione.

---

## Caratteristiche Principali

- **Conversione transazionale**: genera simultaneamente `.htm`, `.rtf`, `.txt` e la cartella asset `<firma>_files`.
- **Modifica assistita in Word**: apertura isolata del documento con rilevamento automatico di salvataggio e chiusura.
- **Normalizzazione HTML & CSS**: pulizia del markup Office superfluo e inlining delle regole CSS per massima compatibilità con i client email.
- **Gestione asset sicura**: percorso immagini relativo con protezione anti path-traversal e configurazione automatica di invio inline su Outlook.
- **Protezione con backup ZIP**: snapshot preventivo prima di qualsiasi sovrascrittura nella cartella firme di Outlook, con ripristino snapshot a uno stato pulito.
- **Supporto multi-architettura**: launcher unificato con rilevamento automatico della bitness di Office (x86/x64).

---

## Requisiti

- **OS**: Windows 10 / 11.
- **Software**: Microsoft Word (necessario per editing e conversione) e Microsoft Outlook Classic (per l'uso diretto delle firme).
- **Sviluppo**: .NET 10.0 SDK e PowerShell.

---

## Quick Start

### Utente Finale
1. Avvia l'applicazione (tramite installer o launcher `OnlyFirmaOutlook.Launcher.exe`).
2. Scegli un modello predefinito (`media/`) o carica un file Word personale.
3. Assegna un nome alla firma e seleziona l'account Outlook di destinazione.
4. Clicca su **Modifica firma**, apporta le modifiche in Word, salva e chiudi Word.
5. Scegli il formato HTML (Filtrato o Completo) e premi **Converti e salva firma**.
6. *Al primo utilizzo*: riavvia Outlook Classic affinché le impostazioni di invio immagini inline abbiano effetto.

### Sviluppatore
```powershell
# Compilazione soluzione
dotnet build OnlyFirmaOutlook.sln -c Release

# Esecuzione test unitari (65 test xUnit)
dotnet test OnlyFirmaOutlook.sln -c Release --no-build

# Build e publish per entrambi i runtime (win-x86 e win-x64)
./scripts/Build-App.ps1 -Configuration Release -PublishMode FrameworkDependent

# Packaging installer EXE (Inno Setup) e MSI (WiX)
./scripts/Package-App.ps1
```

---

## Indice della Documentazione

La documentazione dettagliata è suddivisa per ambito nella cartella [`docs/`](docs/):

| Documento | Ambito e Contenuto |
| :--- | :--- |
| **[Guida Utente](docs/user-guide.md)** | Workflow dettagliato, preset, formati HTML, gestione firme esistenti, backup e ripristino. |
| **[Architettura Tecnica](docs/architecture.md)** | Mappa componenti, rilevamento bitness Office, pipeline transazionale di conversione e vincoli. |
| **[Sviluppo & Packaging](docs/development.md)** | Prerequisiti toolchain, script PowerShell di automazione, CI GitHub Actions e packaging installer. |
| **[Troubleshooting](docs/troubleshooting.md)** | Risoluzione problemi (immagini inline, Word lock, percorsi UNC, gestione dei log). |

---

## Licenza

Software proprietario. Vedere il file [LICENSE](LICENSE).

Copyright (c) 2026 Danny Perondi. Tutti i diritti riservati.
