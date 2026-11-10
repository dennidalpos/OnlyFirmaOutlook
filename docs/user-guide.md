# User guide

[Italian reference](user-guide.it.md)

OnlyFirmaOutlook converts Word documents into Outlook Classic signatures. The interface is Italian; the labels below match the controls on screen.

## Requirements and locations

Use Windows 10/11 with Microsoft Word installed locally for editing and conversion. Outlook Classic is required to use the signatures automatically. Without it, export to a local folder and copy the files later.

The default signature folder is `%APPDATA%\Microsoft\Signatures`. The alternative output is `%USERPROFILE%\Documents\OnlyFirmaOutlook\Output`. Review the selected destination and its writable indicator before conversion.

## Create a signature

1. Start the installed application or `OnlyFirmaOutlook.Launcher.exe`. Choose an existing `.docx`, `.doc` or `.rtf`, or a preset if your distribution supplies one. The source checkout contains only a media placeholder; it does not ship ready-made signature documents.
2. Enter the signature name and select the Outlook account. Invalid filename characters are sanitized. If no account is available, enter an identifier; the final name uses the signature name and identifier.
3. Confirm the output folder. A different folder produces export files without installing them in Outlook.
4. Select **Modifica firma**. The document opens in Word from an isolated local editing copy, including when the source is on a UNC share.
5. Edit, save and close Word. The application observes the saved copy and the Word lock before enabling conversion.
6. Choose filtered HTML for a reduced Office markup result or full HTML for more of Word's original formatting. Neither option guarantees identical rendering in every recipient's mail client.
7. Select **Converti e salva firma** and review the result. The output consists of `<name>.htm`, `.rtf`, `.txt` and `<name>_files/` with relative image references. The destination opens in Explorer.

The conversion engine uses installed Word through COM. It is not a server-side or Office-independent converter.

## Existing signatures and backups

The signature list can remove the selected signature's three formats and associated asset folder. Review the selection before deletion.

Overwriting in the standard Outlook folder first creates a timestamped ZIP snapshot; conversion is refused when the required backup cannot be created. Select a snapshot and **Ripristina backup** to restore it. Restore aligns the signature folder to the snapshot and can remove orphan files; it is a data-changing action. Keep a separate copy before intentionally replacing an existing folder state. Removing obsolete backup ZIPs is also an explicit user action.

## Inline images and Outlook

At startup the application sets the current user's `Send Pictures With Document` option for the supported Office registry branches (16.0, 15.0, 14.0). Restart Outlook Classic after first use, then select the signature in **File → Opzioni → Posta → Firme** for new messages and replies as needed.

A local conversion does not prove that a recipient sees the expected layout. Check the resulting signature in your supported Outlook environment using synthetic content. New Outlook, webmail and third-party clients are outside the implemented Outlook integration.

See [troubleshooting](troubleshooting.md), [architecture](architecture.md) and [development](development.md).
