# Risoluzione Problemi e Note Operative

Diagnostica e soluzioni rapide per anomalie comuni durante l'uso di OnlyFirmaOutlook.

---

## 1. Immagini Assenti o Non Visibili nelle Email Inviate

- **Causa comune**: Outlook Classic non ha ancora caricato la chiave di registro che impone l'invio delle immagini come allegati inline (`Send Pictures With Document`).
- **Soluzione**:
  1. Chiudi completamente Microsoft Outlook.
  2. Avvia OnlyFirmaOutlook (che riscrive la configurazione corretta nel registro utente).
  3. Riapri Outlook Classic e invia un'email di prova.
  4. Assicurati che l'HTML della firma contenga il riferimento relativo `<NomeFirma>_files/immagine.ext` e che la cartella sia presente in `%APPDATA%\Microsoft\Signatures`.

---

## 2. Il Pulsante "Converti e Salva Firma" Rimane Disabilitato

Il pulsante si abilita solo quando sono soddisfatte tutte le seguenti condizioni:
- È stato selezionato un documento sorgente valido.
- Il nome della firma non è vuoto.
- La cartella di destinazione è valida e scrivibile.
- Il documento è stato modificato e salvato in Word.
- **Word è stato chiuso**: se `WINWORD.EXE` mantiene il file bloccato, l'indicatore "Word in esecuzione" resta visibile e la conversione è bloccata per prevenire conflitti.

---

## 3. Microsoft Word Non Rilevato

- **Errore**: all'avvio compare un avviso di "Word non trovato".
- **Causa**: l'app richiede Microsoft Word locale per avviare il motore COM interop di conversione (`Word.Application`).
- **Soluzione**: verificare che Office/Word sia regolarmente installato e associato all'utente Windows corrente.

---

## 4. Outlook Non Installato o Senza Account

- Se Outlook Classic non è installato o non possiede account configurati:
  1. L'app propone una cartella alternativa in `%USERPROFILE%\Documents\OnlyFirmaOutlook\Output`.
  2. Compare un campo per specificare manualmente un identificativo per la firma.
  3. I file generati (`.htm`, `.rtf`, `.txt` e cartella asset) possono essere copiati manualmente in seguito.

---

## 5. File di Rete (Percorsi UNC)

- Se si apre un file sorgente situato su una share di rete (es. `\\server\share\firma.docx`), OnlyFirmaOutlook ne crea automaticamente una copia temporanea locale prima di avviare l'editing in Word, prevenendo lock di rete e latenze di salvataggio.

---

## 6. Consultazione e Pulizia Log

- **Pannello UI**: mostra i messaggi operativi in tempo reale.
- **Pulsante "Copia log"**: copia l'intero buffer negli appunti.
- **Pulsante "Apri file log"**: apre il file su disco con Blocco Note.
- **Percorso log su disco**: `%LOCALAPPDATA%\OnlyFirmaOutlook\Logs\app.log`
- **Pulsante "Pulisci log"**: svuota la schermata e cancella il file `app.log` su disco.

---

## 7. Pulizia File Temporanei

- OnlyFirmaOutlook ripulisce automaticamente le proprie cartelle di sessione ed editing alla chiusura.
- In caso di chiusura forzata o anomala del sistema, le cartelle orfane con più di 24 ore vengono epurate al successivo avvio.
- Percorsi temporanei di lavoro:
  - `%LOCALAPPDATA%\OnlyFirmaOutlook\Temp\`
  - `%LOCALAPPDATA%\OnlyFirmaOutlook\EditorTemp\`
