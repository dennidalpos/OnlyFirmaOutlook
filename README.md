# OnlyFirmaOutlook

WPF desktop app (.NET 10) for Windows that converts Microsoft Word documents (`.docx`, `.doc`, `.rtf`) into ready-to-use Outlook Classic signatures, with HTML normalization, local image asset handling, automatic backups and installer packaging.

## Features

- **Transactional conversion**: generates `.htm`, `.rtf`, `.txt` and the `<signature>_files` asset folder together.
- **Assisted editing in Word**: opens the document in an isolated Word instance and detects save and close automatically.
- **HTML and CSS normalization**: strips superfluous Office markup and inlines CSS rules for better email client compatibility.
- **Safe asset handling**: relative image paths with path-traversal protection, and automatic configuration of inline image sending in Outlook.
- **ZIP backup protection**: a snapshot is taken before any overwrite in the Outlook signatures folder, with snapshot restore.
- **x86 and x64 Office support**: a unified launcher detects the bitness of the installed Office.

## Requirements

- Windows 10 or 11.
- Microsoft Word (needed for editing and conversion) and Microsoft Outlook Classic (to use the signatures).
- To build: .NET 10.0 SDK and PowerShell.

## Usage

1. Start the app (installer or `OnlyFirmaOutlook.Launcher.exe`).
2. Pick a preset from `media/` or load your own Word file.
3. Name the signature and select the target Outlook account.
4. Click **Modifica firma**, edit in Word, save and close Word.
5. Choose the HTML format (filtered or full) and click **Converti e salva firma**.
6. On first use, restart Outlook Classic so the inline image setting takes effect.

The user interface is in Italian.

## Build

```powershell
dotnet build OnlyFirmaOutlook.sln -c Release
dotnet test OnlyFirmaOutlook.sln -c Release --no-build

# Build and publish win-x86 and win-x64
./scripts/Build-App.ps1 -Configuration Release -PublishMode FrameworkDependent

# Installer packaging: EXE (Inno Setup) and MSI (WiX)
./scripts/Package-App.ps1
```

## Documentation

Documentation is written in Italian, in [`docs/`](docs/):

| Document | Content |
| :--- | :--- |
| [User guide](docs/user-guide.md) | Workflow, presets, HTML formats, existing signatures, backup and restore |
| [Architecture](docs/architecture.md) | Components, Office bitness detection, transactional conversion pipeline |
| [Development and packaging](docs/development.md) | Toolchain, automation scripts, CI and installer packaging |
| [Troubleshooting](docs/troubleshooting.md) | Inline images, Word locks, UNC paths, logs |

## Security

The app works on local files and the current user's Outlook signatures folder; it has no network component. See [SECURITY.md](SECURITY.md) to report a vulnerability.

## License

Proprietary, all rights reserved. See [LICENSE](LICENSE). Copyright (c) 2026 Danny Perondi.
