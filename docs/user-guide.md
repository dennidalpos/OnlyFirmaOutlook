# Guida Utente

OnlyFirmaOutlook trasforma documenti Word (.docx, .doc, .rtf) in firme per Microsoft Outlook Classic con supporto a preset, modifica assistita, asset locali e backup automatici.

---

## Requisiti

- **Sistema operativo**: Windows 10 / 11.
- **Microsoft Word**: installato localmente per l'editing e la conversione.
- **Microsoft Outlook Classic**: per l'uso automatico delle firme (se assente, è possibile esportare in una cartella locale).

---

## Percorsi Predefiniti

- **Cartella firme Outlook Classic**: `%APPDATA%\Microsoft\Signatures`
- **Destinazione alternativa**: `%USERPROFILE%\Documents\OnlyFirmaOutlook\Output`

---

## Flusso Operativo Passo-Passo

### 1. Selezione del Documento
- **Preset**: modelli pronti all'uso distribuiti con l'app (cartella `media`).
- **File personalizzato**: supporta `.docx`, `.doc` e `.rtf`. I file provenienti da percorsi di rete (UNC) vengono copiati automaticamente in locale per evitare lock o rallentamenti durante la modifica.

### 2. Nome Firma e Account
- **Nome firma**: viene ripulito automaticamente da caratteri non validi per il file system.
- **Account**: se Outlook è configurato, seleziona l'account a cui associare la firma. In assenza di account, inserisci un identificativo (es. email o nome utente): verrà aggiunto al nome finale (`<Nome> - <Identificativo>`).

### 3. Cartella di Destinazione
- La cartella predefinita è quella di Outlook (`%APPDATA%\Microsoft\Signatures`).
- L'indicatore verifica in tempo reale se la cartella è scrivibile.
- Se Outlook non è disponibile o si sceglie un'altra cartella, le firme possono essere archiviate o copiate manualmente in seguito.

### 4. Modifica in Word
1. Clicca su **Modifica firma**: il documento viene aperto in Word da una cartella di lavoro temporanea isolata.
2. Effettua le modifiche desiderate, salva (`Ctrl+S` o `Maiusc+F12`) e chiudi Word.
3. L'app monitora lo stato del file: la conversione si abilita solo quando il documento risulta salvato e Word è chiuso.

### 5. Formato HTML
- **HTML Filtrato**: rimuove gli stili Microsoft superflui, riduce il peso e massimizza la compatibilità con i client email riceventi.
- **HTML Completo**: conserva una maggiore fedeltà visiva rispetto al layout Word originale.

### 6. Conversione e Pubblicazione
1. Clicca su **Converti e salva firma**.
2. L'app genera i tre formati richiesti da Outlook Classic:
   - `<Nome>.htm` (con riferimenti relativi alle immagini)
   - `<Nome>.rtf` (per messaggi RTF)
   - `<Nome>.txt` (per messaggi in testo semplice)
   - `<Nome>_files/` (cartella contenente gli asset grafici)
3. Al termine, la cartella di destinazione viene aperta automaticamente in Esplora Risorse.

---

## Gestione Firme e Sovrascrittura

- **Firme esistenti**: la lista mostra le firme presenti nella cartella di destinazione, con possibilità di eliminazione completa (file `.htm`, `.rtf`, `.txt` e cartella asset associata).
- **Sovrascrittura con backup obbligatorio**: se si sovrascrive una firma nella cartella predefinita di Outlook, l'app genera preventivamente un archivio ZIP (`backup_<timestamp>.zip`). La sovrascrittura procede solo se il backup è andato a buon fine.

---

## Backup e Ripristino

- **Elenco backup**: accessibile nella sezione dedicata; mostra data e ora degli snapshot ZIP disponibili.
- **Ripristino**: seleziona uno snapshot e clicca su **Ripristina backup**. La cartella firme viene riallineata allo stato del backup, eliminando eventuali artefatti orfani.
- **Eliminazione**: consente di rimuovere gli archivi ZIP obsoleti.

---

## Configurazione Outlook e Immagini Inline

Al primo avvio, OnlyFirmaOutlook configura nel registro utente l'opzione `Send Pictures With Document` per le versioni supportate di Outlook (16.0, 15.0, 14.0).

> [!IMPORTANT]
> **Riavvio di Outlook**: chiudi e riavvia Outlook Classic dopo aver eseguito OnlyFirmaOutlook affinché l'opzione di invio immagini inline diventi attiva.
>
> In Outlook, vai in **File → Opzioni → Posta → Firme** e imposta la nuova firma per i messaggi desiderati (nuovi messaggi / risposte).
