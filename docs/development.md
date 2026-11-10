# Development and packaging

[Italian reference](development.it.md)

Use Windows with the .NET 10 SDK and PowerShell. Functional document conversion also requires installed Word. Installer generation requires Inno Setup 6 for EXE and WiX v4+ for MSI; inspect the scripts before provisioning either tool.

## Build and test

From the repository root:

~~~powershell
dotnet build OnlyFirmaOutlook.sln -c Release
dotnet test OnlyFirmaOutlook.sln -c Release --no-build
~~~

Use the test output from the current revision; historical test counts are not a fixed contract. Automated unit tests do not qualify Word COM, Outlook configuration, installer effects or a clean Windows host.

Managed verification output belongs in an initialized SecondBrain runtime workspace. The public scripts use their repository-local output layout; an environment variable alone does not redirect every output. Run in an isolated source copy or verify explicit SDK artifact/intermediate paths before using them.

## Publish

Build-App publishes win-x86 and win-x64, includes the launcher and copies media presets into dist:

~~~powershell
./scripts/Build-App.ps1 -Configuration Release -PublishMode FrameworkDependent
./scripts/Build-App.ps1 -Configuration Release -PublishMode SelfContained
~~~

Build-App performs its default clean unless SkipClean is specified; preserve referenced outputs before intentionally running that path. The media source currently contains only a placeholder; add reviewed distributable templates if they are part of the intended package.

Without PublishMode, the script requests the publication mode interactively. SkipTests deliberately skips tests and therefore cannot serve as test evidence. Framework-dependent output requires an appropriate installed runtime; self-contained output includes runtime files.

## Installer toolchain and packages

Install-InstallerToolchain without Install inspects toolchain availability. Its Install switch provisions software and is a separate, intentional machine-changing operation.

~~~powershell
./scripts/Install-InstallerToolchain.ps1
./scripts/Package-App.ps1
./scripts/Package-App.ps1 -Format InnoSetup
./scripts/Package-App.ps1 -Format WiX -SkipBuild
~~~

Package-App builds when necessary and writes installer output under packaging/output. SkipBuild reuses existing dist files, so verify their revision and architecture before packaging. Do not infer a new product version from a documentation change.

The installers are per-user distributions under LocalAppData. An alternative deployment copies the unified dist folder to a managed share and starts OnlyFirmaOutlook.Launcher.exe. A generated installer or compiled executable is not evidence of a signed, installed or published release.

## CI and qualification

The versioned workflow describes .NET 10 restore/build/tests and publication-artifact checks. Repository Actions availability is an independent remote setting; the maintainer's recorded workflow uses local verification. Inspect the current workflow and GitHub setting rather than claiming a new CI run.

Before release, distinguish unit-test results, functional Word/Outlook tests, bitness selection, clean-host installation, signing and publication. Never use real customer signatures in public fixtures.

## Cleanup

Clean-App and its All/IncludeUserData options delete generated or user data. They are not a prerequisite for ordinary documentation work. Inspect the concrete scope and preserve referenced output, backups and fixtures; do not run cleanup without the required explicit request.

See [architecture](architecture.md), [user guide](user-guide.md), [troubleshooting](troubleshooting.md).
