# Troubleshooting

[Italian reference](troubleshooting.it.md)

Use synthetic documents when reproducing a problem. Preserve relevant logs, signature files and backups before any operation that replaces or deletes them.

## Missing inline images

Close Outlook Classic, start OnlyFirmaOutlook to apply its current-user inline-image setting, then reopen Outlook and send a synthetic test message. Check that HTML references `<signature>_files/image.ext` relatively and that the asset folder is present beside the .htm file.

The registry setting does not guarantee rendering by every recipient's email client.

## Conversion stays disabled

Check the source document, nonempty signature name and writable destination. Use **Modifica firma**, save the isolated Word document and close its Word instance. A remaining file lock or **Word in esecuzione** state prevents conversion. Do not force-stop unrelated Word processes or discard unsaved work to clear the indicator.

## Word unavailable

Installed Microsoft Word is required by the implemented `Word.Application` COM converter. Verify the current user's Office installation and registration. The app is not an Office-independent OpenXML converter.

## Outlook unavailable

Without Outlook Classic/accounts the app offers a local output folder and an identifier field. Generated HTML/RTF/text/assets can be copied later. New Outlook does not provide the Classic signature integration used here.

## UNC input

A document selected on a share is copied to a local editing workspace. Verify access to the original and the local copy, especially after interrupted network access; successful staging does not qualify all network-share configurations.

## Logs

The UI shows operational messages. **Copia log** copies the buffer; **Apri file log** opens the disk log at `%LOCALAPPDATA%\OnlyFirmaOutlook\Logs\app.log`. **Pulisci log** clears the view and deletes the disk log. Capture a sanitized diagnostic copy before intentionally using it; do not publish raw personal paths or document content.

## Temporary folders

The application cleans its own editing/conversion sessions and removes eligible orphan folders older than 24 hours on a later start. Their roots are LocalAppData/OnlyFirmaOutlook/Temp and EditorTemp. Preserve an interrupted reproduction before restarting when those files are needed as evidence. Do not manually delete other sessions' folders.

See [user guide](user-guide.md) and [security reporting](../SECURITY.md).
