# Outlook/Word signature table borders — research notes (offline)

## Internet access attempts

While preparing this analysis, external web access was blocked by the environment (HTTP 403 from an Envoy proxy) for both HTTP and HTTPS requests. This prevented collecting *latest news* and direct citations from Microsoft or forum sources.

Commands attempted (representative):

- `curl -I -s http://example.com` → `HTTP/1.1 403 Forbidden`
- `curl -I -L -s -A 'Mozilla/5.0' 'https://www.bing.com/search?q=Outlook+signature+table+borders+Word+HTML+export'` → `HTTP/1.1 403 Forbidden`

## Known technical pattern (general knowledge)

Even without external access, the issue described is consistent with a well‑known Outlook/Word HTML export behavior:

- Outlook uses the Word HTML rendering engine. Word exports tables with Office‑specific CSS (e.g., `class=MsoTableGrid`, `mso-border-alt`, `mso-border-top-alt`, etc.).
- In many cases, Word inserts `border: solid` or `mso-border-alt: solid` values in the exported HTML, even if table borders were hidden in Word.
- Outlook’s Word-based renderer can ignore or override some “visual” Word settings when it converts to HTML, causing hidden borders to reappear in the final signature.

## Practical remediation steps (aligned with the user’s draft)

These are the most reliable and repeatable HTML fixes when Word-exported signatures show visible table borders:

1. **Export as “Pagina Web, filtrata”** to reduce Office-specific tags.
2. **Remove `MsoTableGrid` styling**, since it tends to reintroduce borders.
3. **Force borders off**:
   - Replace `border:solid` → `border:none`.
   - Replace `mso-border-alt:solid` → `mso-border-alt:none`.
4. **Ensure table attributes explicitly disable borders**:
   - `<table border="0" cellspacing="0" cellpadding="0">`
5. **Normalize cells** (optional): add `style="border:none"` to `<td>` or `<th>`.
6. **Verify with Outlook rendering**, not just a browser, because Outlook uses the Word engine.

## Recommended follow‑up if web access becomes available

- Search Microsoft Answers / Tech Community and Outlook forums for “signature table borders Word HTML export” and “MsoTableGrid border” and collect recent reports (last 12 months).
- Check Office release notes for Word or Outlook HTML rendering regressions.
- Compare a “clean” HTML export with the problematic one to isolate specific CSS/attributes being reintroduced.
