# OnlyFirmaOutlook

Applicazione Windows per convertire documenti Word in firme email Outlook.

## Funzionalità

- Conversione documenti Word (.doc/.docx) e RTF in firme Outlook (HTML, RTF, TXT)
- Immagini embedded base64 per compatibilità universale con tutti i provider email
- Fix bug Outlook 2512+ (bordi tabelle indesiderati)
- Gestione preset di firma preconfigurati
- Backup e ripristino firme esistenti
- Rilevamento account Outlook (Exchange, IMAP, POP3, deleghe)
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

## Utilizzo

1. **Seleziona preset** o carica un documento Word/RTF personalizzato
2. **Inserisci nome firma** che apparirà in Outlook
3. **Seleziona cartella destinazione** (default: cartella firme Outlook)
4. **Seleziona account** Outlook (opzionale, per suffisso nel nome firma)
5. **Configura opzioni export**:
   - Formato HTML: Filtrato o Completo
   - Fix bordi tabelle Outlook 2512+ (attivo di default)
6. **Converti e salva** la firma

## Opzioni export

### Formato HTML
- **HTML Filtrato**: più leggero, rimuove metadati Word
- **HTML Completo**: preserva tutti gli stili Word (consigliato)

### Fix bordi tabelle Outlook 2512+
Corregge il bug introdotto in Outlook Classic versione 2512 che aggiunge bordi indesiderati alle tabelle nelle firme.

**Quando attivarlo**: sempre, se usi Outlook Classic aggiornato
**Quando disattivarlo**: se noti problemi con tabelle che devono avere bordi visibili

## Struttura progetto

```
OnlyFirmaOutlook/
├── src/
│   ├── Bootstrapper/          # Launcher 32/64-bit
│   ├── OnlyFirmaOutlook/      # App WPF principale
│   │   ├── Models/
│   │   ├── Services/
│   │   ├── Views/
│   │   └── media/             # Preset firma
│   └── Shared/
├── tests/
│   └── OnlyFirmaOutlook.Tests/
├── scripts/
│   ├── build.ps1
│   └── clean.ps1
└── dist/                      # Output build
```

## Comandi

### Build
```powershell
.\scripts\build.ps1
```

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
| `<app>\media\` | Preset firma (.doc/.docx/.rtf) |

## Distribuzione

1. Eseguire `.\scripts\build.ps1`
2. Copiare `dist\` su share di rete o PC target
3. Eseguire `OnlyFirmaOutlook.Launcher.exe`

## Preset firma

Per aggiungere preset:
1. Creare documento Word/RTF con layout firma
2. Copiare in `src/OnlyFirmaOutlook/media/`
3. Ricompilare

## Tecnologie

- .NET 8 / WPF
- HtmlAgilityPack
- PreMailer.Net
- Microsoft.Office.Interop.Word

## Licenza

MIT
