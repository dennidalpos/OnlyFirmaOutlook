# Sviluppo, Automazione e Packaging

Guida per sviluppatori, DevOps e manutentori del repository OnlyFirmaOutlook (.NET 10).

---

## Prerequisiti di Sviluppo

- **.NET 10.0 SDK** (target `net10.0-windows`).
- **PowerShell 7+** (o Windows PowerShell 5.1).
- **Microsoft Word** locale (per test funzionali di conversione).
- **Toolchain di packaging** (opzionale, solo per creare installer):
  - **Inno Setup 6** (`ISCC.exe`) per installer EXE.
  - **WiX Toolset v4+** (`wix.exe`) per pacchetti MSI.

---

## Bootstrap Toolchain Packaging

Lo script `Install-InstallerToolchain.ps1` controlla e installa automaticamente i prerequisiti mancanti via `winget` e `dotnet tool`:

```powershell
# Verifica stato dei componenti
./scripts/Install-InstallerToolchain.ps1

# Installazione automatica prerequisiti mancanti
./scripts/Install-InstallerToolchain.ps1 -Install
```

---

## Compilazione e Test Manuali

```powershell
# Compilazione soluzione in Release
dotnet build OnlyFirmaOutlook.sln -c Release

# Esecuzione completa dei test xUnit (65 test)
dotnet test OnlyFirmaOutlook.sln -c Release --no-build
```

---

## Script di Automazione PowerShell

Gli script seguono la convenzione standard `Verb-Noun.ps1` (sono disponibili anche i wrapper `build.ps1` e `clean.ps1`).

### 1. Build e Publish (`Build-App.ps1`)
Compila e pubblica l'app per entrambi i runtime (`win-x86` e `win-x64`), include il bootstrapper e copia i preset `media/` nella cartella `dist/`:

```powershell
# Build e publish Release (richiede interattivamente la modalità di publish)
./scripts/Build-App.ps1 -Configuration Release

# Publish forzando Framework-dependent (.NET runtime condiviso)
./scripts/Build-App.ps1 -Configuration Release -PublishMode FrameworkDependent

# Publish forzando Self-contained (runtime incluso)
./scripts/Build-App.ps1 -Configuration Release -PublishMode SelfContained

# Compilazione rapida saltando i test
./scripts/Build-App.ps1 -Configuration Release -SkipTests
```

### 2. Packaging Installer (`Package-App.ps1`)
Esegue la build (se non già presente) e genera gli installer in `packaging/output/`:

```powershell
# Genera sia EXE (Inno Setup) che MSI (WiX)
./scripts/Package-App.ps1

# Solo installer Inno Setup EXE con versione custom
./scripts/Package-App.ps1 -Format InnoSetup -Version "1.0.0"

# Solo installer WiX MSI riutilizzando i binari dist/ esistenti
./scripts/Package-App.ps1 -Format WiX -SkipBuild
```

### 3. Pulizia Artefatti (`Clean-App.ps1`)
```powershell
# Pulizia standard (bin, obj, dist, packaging/output, TestResults)
./scripts/Clean-App.ps1

# Pulizia completa (include .vs e directory temporanee utente)
./scripts/Clean-App.ps1 -All -IncludeUserData
```

---

## Pipeline di Continuous Integration (GitHub Actions)

Il workflow in `.github/workflows/ci.yml` gira su runner `windows-latest` ed esegue:
1. Setup di .NET 10.
2. `dotnet restore` e `dotnet build -c Release`.
3. `dotnet test -c Release --no-build`.
4. Verifica funzionale di `Build-App.ps1` con generazione e validazione di artefatti `dist/` (Launcher, binari `win-x86`/`win-x64` e preset di test).

---

## Modalità di Distribuzione

1. **Installer Standalone (Scelta consigliata)**:
   - Distribuire `OnlyFirmaOutlook-Setup-<versione>.exe` (Inno Setup) oppure `OnlyFirmaOutlook-<versione>.msi` (WiX) prodotti da `Package-App.ps1`.
   - Installano l'app per singolo utente in `%LOCALAPPDATA%\OnlyFirmaOutlook`.
2. **Share di Rete Centralizzata**:
   - Copiare il contenuto della cartella `dist/` su una cartella condivisa di rete.
   - Gli utenti avviano `OnlyFirmaOutlook.Launcher.exe`, che rileva l'architettura di Office ed esegue la versione corretta (`win-x86` o `win-x64`).
