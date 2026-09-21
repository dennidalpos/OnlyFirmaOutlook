# AGENTS.md

`v1.2 · 2026-09-21` — Non-derivable repository facts only. Cap ~2500 characters.

## 1. Identity & Scope
- **Purpose**: App desktop WPF (.NET 10) per Windows che converte documenti Word (.doc/.docx/.rtf) in firme Outlook Classic con normalizzazione HTML, asset locali, backup ZIP e packaging per installazione.
- **Runtime / Toolchain**: .NET 10.0 SDK (`net10.0-windows`), PowerShell 5.1/7+, Inno Setup 6 (`ISCC.exe`), WiX Toolset v7 (`wix.exe`).
- **Out of Scope**: OS non-Windows, client email non-Outlook Classic, sync cloud/backend.
- **Hard Constraints**: Richiede Windows. Le immagini delle firme usano path relativi `<firma>_files`. Bitness selezionata all'avvio dal Launcher in base a Office.

## 2. Verified Commands
Update date on verification.

| Workflow | Command | Shell / Cwd | Verified on | Notes / Examples |
| :--- | :--- | :--- | :--- | :--- |
| **Fast Verification** | `dotnet test OnlyFirmaOutlook.sln` | pwsh / root | 2026-09-21 | 65 unit test xUnit, zero warning |
| **Full Build** | `dotnet build OnlyFirmaOutlook.sln -c Release` | pwsh / root | 2026-09-21 | Compilazione Release |
| **Build + Publish** | `./scripts/Build-App.ps1 -Configuration Release -PublishMode FrameworkDependent` | pwsh / root | 2026-09-21 | Genera dist/ con win-x86 e win-x64 |
| **Package All** | `./scripts/Package-App.ps1` | pwsh / root | 2026-09-21 | Genera EXE (Inno Setup) e MSI (WiX) in packaging/output/ |
| **Toolchain Check**| `./scripts/Install-InstallerToolchain.ps1` | pwsh / root | 2026-09-21 | Verifica .NET 10, winget, Inno Setup, WiX |
| **Clean** | `./scripts/Clean-App.ps1` | pwsh / root | 2026-09-21 | Pulisce bin, obj, dist, packaging/output |

## 3. Architecture & Boundaries
- `src/OnlyFirmaOutlook`: App WPF (`MainWindow` divisa in partial class: Editor, SignatureManagement, Chrome).
- `src/Bootstrapper`: Launcher x64/x86 basato su `OfficeBitnessDetector`.
- `packaging/innosetup/setup.iss`: Configurazione Inno Setup per setup EXE canonico.
- `packaging/msi/OnlyFirmaOutlook.wxs`: Sorgente WiX per pacchetto MSI canonico.
- `scripts/`: Script PowerShell con convenzione `Verb-Noun.ps1` (compatibilità con `build.ps1` e `clean.ps1`).
- `docs/`: Documentazione per ambito (`user-guide.md`, `architecture.md`, `development.md`, `troubleshooting.md`).

## 4. Sensitive Areas & Gotchas
- **WiX v7 OSMF**: La compilazione richiede il flag `--acceptEula wix7`.
- **Word COM Interop**: `WordConversionService` dipende da Microsoft Word locale.
- **Outlook Inline Images**: L'app imposta `Send Pictures With Document` nel registro dell'utente corrente; Outlook va riavviato al primo avvio.
