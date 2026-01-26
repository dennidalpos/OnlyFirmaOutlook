# Linee guida per firme universali (Outlook Classic + client esterni)

Queste indicazioni aiutano a creare firme stabili e compatibili con Outlook Classic, webmail e client esterni. L’obiettivo è ridurre al minimo i problemi di rendering, soprattutto nei client con supporto HTML limitato.

## 1) Struttura HTML consigliata

- Usa una struttura semplice con `table` per il layout principale.
- Evita layout complessi basati su `div` annidati o `flex`, non sempre supportati.
- Imposta larghezze fisse o massime in pixel per evitare ridimensionamenti imprevedibili.

Esempio base:

```html
<table cellpadding="0" cellspacing="0" border="0" style="font-family:Calibri, Arial, sans-serif; font-size:12px; line-height:1.3;">
  <tr>
    <td style="padding:0; vertical-align:top;">
      <strong>Nome Cognome</strong><br>
      Ruolo · Azienda<br>
      +39 000 000000 · nome@azienda.it
    </td>
  </tr>
</table>
```

## 2) Stili: solo inline

- Usa esclusivamente **stili inline** (`style=""`), evita `<style>` e classi CSS.
- Mantieni stili semplici: `font-family`, `font-size`, `color`, `line-height`, `text-decoration`, `padding`, `margin`.
- Evita proprietà avanzate come `float`, `position`, `z-index`, `flex`, `grid`, `background-image`.

## 3) Font e dimensioni

- Scegli font comuni: `Calibri`, `Arial`, `Helvetica`, `Verdana`.
- Mantieni dimensioni tra **11–13px** per il testo principale.
- Usa `line-height: 1.3` o `1.4` per leggibilità.

## 4) Immagini

- Usa formati **PNG/JPG** con dimensioni ottimizzate.
- Limita larghezza a **150–300px** per loghi.
- Evita immagini troppo pesanti (consigliato < 200 KB).
- Inserisci `alt` per le immagini.

Esempio:

```html
<img src="logo.png" alt="Azienda" width="180" style="display:block; border:0;">
```

## 5) Colori e contrasto

- Usa colori con buon contrasto (testo scuro su sfondo chiaro).
- Evita testi molto chiari o su sfondo con immagini.
- Limita l’uso di colori diversi: massimo 2–3 colori principali.

## 6) Link

- Sempre in formato completo: `https://`.
- Evita link troppo lunghi visivamente, ma mantieni l’URL completo nel tag `href`.

Esempio:

```html
<a href="https://www.azienda.it" style="color:#0078D4; text-decoration:none;">www.azienda.it</a>
```

## 7) Spaziatura e separatori

- Usa `padding` e `line-height` per gestire la spaziatura, evitando `margin` eccessivi.
- I separatori verticali o orizzontali sono più affidabili se realizzati con celle `td` e bordi sottili.

Esempio separatore verticale:

```html
<td style="border-left:1px solid #CCCCCC; padding-left:12px;"></td>
```

## 8) Compatibilità con client esterni

Per aumentare la compatibilità:

- Evita HTML complesso e JavaScript.
- Non usare form o elementi interattivi.
- Mantieni il contenuto lineare e leggibile anche senza stili avanzati.
- Testa con almeno: Outlook Classic, Outlook Web, Gmail Web, Apple Mail.

## 9) Firma “fallback” testuale

È utile includere anche una versione testuale coerente:

```
Nome Cognome
Ruolo · Azienda
+39 000 000000 · nome@azienda.it
www.azienda.it
```

## 10) Verifica finale

Prima dell’uso:

- Invia una mail di prova verso client esterni.
- Controlla che i link funzionino e le immagini siano visibili.
- Verifica che il layout non si “rompa” su schermi piccoli.

## Suggerimento pratico

Se parti da Word, mantieni il documento semplice e lineare, poi usa l’opzione HTML filtrato se noti stili eccessivi o layout instabile.
