# OnlyFirmaOutlook

OnlyFirmaOutlook è un'applicazione desktop Windows (WPF) che trasforma documenti Word in firme compatibili con Outlook Classic. L'app guida l'utente nella modifica del documento, nella scelta del formato HTML e nella gestione delle firme esistenti, includendo anche la creazione e il ripristino dei backup.

## Cosa fa

- Converte file Word (.doc/.docx) in firme Outlook (HTML/RTF/TXT + cartelle assets).
- Integra l'apertura del documento in Word per una modifica obbligatoria prima della conversione.
- Carica preset da una cartella `media` ed accetta file personalizzati.
- Rileva account Outlook se presenti, altrimenti usa un identificativo manuale.
- Gestisce firme esistenti con avvisi di sovrascrittura ed eliminazione dedicata.
- Crea backup ZIP nella cartella firme e permette ripristino/eliminazione.
- Supporta build x86/x64 con launcher che seleziona la bitness corretta in base a Office.
- Consente l'esecuzione da share di rete con copia locale dei file temporanei.

## Requisiti

- Windows 10/11
- Microsoft Word 2013+ (2013, 2016, 2019, 2021, 365)
- .NET 8.0 Runtime (incluso nelle build self-contained)
- Microsoft Outlook (opzionale)

## Avvio rapido (utenti finali)

1. Avvia `OnlyFirmaOutlook.Launcher.exe` dalla cartella di distribuzione.
2. Seleziona un preset o carica un documento Word.
3. Modifica il documento in Word, salva (Maiusc+F12) e chiudi.
4. Imposta il nome della firma (e l'account se disponibile).
5. Scegli il formato HTML e conferma la conversione.
6. Se necessario, ripristina un backup o elimina quelli obsoleti.

## Backup firme

- I backup vengono creati automaticamente prima della conversione.
- I file sono salvati nella cartella firme di Outlook in formato ZIP con prefisso:
  - `backup_firme_onlyfirmaoutlook_yyyy-MM-dd-HH-mm.zip`
- Dal punto 7 dell'interfaccia puoi:
  - Visualizzare i backup presenti.
  - Ripristinare un backup (sovrascrive i file correnti nella cartella firme).
  - Eliminare backup non più necessari.

## Preset

I preset sono documenti Word disponibili nella cartella `media` dell'app:

- Percorso runtime: `AppContext.BaseDirectory\media`
- I file temporanei `~$` di Word vengono ignorati.

Per includerli nella distribuzione, copia i file in `src/OnlyFirmaOutlook/media` e usa lo script di build.

## Immagini inline (CID) nelle firme HTML

Se vuoi inviare firme con immagini realmente embedded (Content-ID) quando componi email via Outlook Interop, puoi usare il post-processing HTML incluso nel progetto. Esempio minimo end-to-end:

```csharp
using System.Text;
using Microsoft.Office.Interop.Outlook;
using OnlyFirmaOutlook.Services;

// htmlPath è il file HTML esportato da Word.
var htmlPath = Path.Combine(
    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
    "Microsoft",
    "Signatures",
    "NomeFirma.htm");

var html = File.ReadAllText(htmlPath, Encoding.GetEncoding(1252));
var baseDir = Path.GetDirectoryName(htmlPath)!;
var (htmlRewritten, images) = WordHtmlCidPostProcessor.RewriteLocalImageRefsToCid(html, baseDir);

var outlook = new Application();
var mail = (MailItem)outlook.CreateItem(OlItemType.olMailItem);
mail.BodyFormat = OlBodyFormat.olFormatHTML;
mail.HTMLBody = htmlRewritten;

OutlookCidAttacher.AddInlineCidAttachments(mail, images);
mail.Send();
```

Se vuoi un automatismo pronto all'uso, puoi delegare tutto a:

```csharp
var outlook = new Application();
var mail = (MailItem)outlook.CreateItem(OlItemType.olMailItem);
OutlookSignatureEmbedder.ApplySignatureWithInlineImages(mail, htmlPath);
mail.Send();
```

### Note operative

- Per verificare che le immagini siano realmente inline, controlla il sorgente della mail ricevuta: deve contenere `multipart/related`, header `Content-ID` sugli allegati e riferimenti `src="cid:..."` nell'HTML. La semplice visualizzazione corretta in Outlook non è sufficiente a garantire che l'email inviata abbia gli allegati inline corretti.
- In alternativa, se cambi approccio, puoi inserire le immagini direttamente nel WordEditor di Outlook come `InlineShapes.AddPicture(...)` invece di usare HTML con percorsi locali; in quel caso Outlook gestisce nativamente gli allegati inline, ma richiede un template/placeholder nel documento.

## Struttura del progetto

```
OnlyFirmaOutlook.sln
scripts/
  build.ps1
  clean.ps1
src/
  Bootstrapper/        # Launcher per bitness Office
  OnlyFirmaOutlook/    # App WPF principale
```

## Build e distribuzione

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

Distribuzione consigliata:

- Copiare l'intera cartella `dist` su una share di rete o in una cartella locale.
- L'utente finale avvia sempre `OnlyFirmaOutlook.Launcher.exe`.

## Log e diagnostica

I log sono salvati in:

```
%LOCALAPPDATA%\OnlyFirmaOutlook\Logs\app.log
```

Dall'interfaccia è possibile copiare o pulire il log e aprire il file corrente.

## Licenza

MIT.
