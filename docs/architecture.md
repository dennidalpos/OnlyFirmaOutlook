# Architecture

[Italian reference](architecture.it.md)

OnlyFirmaOutlook is a Windows WPF application targeting `net10.0-windows`. Its conversion engine currently automates locally installed Word through COM.

## Projects and UI

| Path | Responsibility |
|---|---|
| `src/OnlyFirmaOutlook/` | WPF application, models, services, styles and views |
| `src/Bootstrapper/` | Unified launcher selecting the Office architecture |
| `src/Shared/` | Shared Office bitness detection |
| `packaging/innosetup/` | Inno Setup EXE source |
| `packaging/msi/` | WiX MSI source |

MainWindow is split into partial classes for lifecycle/account detection, editing and file watching, signature/backup management, and window chrome/log/guide controls. The partial classes share window state; they do not constitute a completed MVVM separation.

## Launcher

The launcher examines Outlook Bitness registry values for Office 16/15/14, Click-to-Run Platform, MSI installation roots and the Word executable's PE header. When it cannot obtain a conclusive architecture it falls back to win-x64. It starts the win-x86 or win-x64 child executable and forwards arguments.

The fallback is an implementation default, not proof that the detected Office environment is compatible.

## Conversion workflow

~~~mermaid
flowchart LR
    A[Source document] --> B[Isolated editor copy]
    B --> C[Installed Word COM conversion]
    C --> D[HTML normalization and CSS inlining]
    D --> E[Confined local asset copy]
    E --> F[Publish formats with backup and rollback handling]
~~~

WordEditorService works in an isolated editor folder under LocalAppData. Word conversion produces filtered/full HTML, RTF, plain text and an asset directory in a staging area. WordHtmlSignatureNormalizer removes nonrendering/unsafe nodes and surplus Office markup while preserving required visual styles, including mso-line-height-rule. CssInliner applies embedded CSS rules inline.

AssetManager resolves local images within the permitted base directory, checks traversal boundaries and copies assets to the canonical signature folder with relative references. Do not treat the output as a general-purpose sanitizer for arbitrary web content.

Publishing moves the generated formats into the destination with rollback handling. Existing Outlook signatures require a successful ZIP backup before overwrite. SignatureRepository provides listing, deletion, snapshots and restore; restore may remove artifacts absent from the chosen snapshot.

## Outlook and temporary data

OutlookInlineImageSettings writes `Send Pictures With Document=1` beneath the current user's Outlook Mail options for the supported Office versions. Temporary editing/conversion directories live under LocalAppData; cleanup helpers handle transient locks and retained orphan sessions according to the implemented age policy.

Source inputs, signatures, backups and logs can contain personal information. The application has no cloud conversion service, but a selected source/destination can be a network path.

## Boundaries

Windows APIs, WPF, registry and Word COM are runtime dependencies. Outlook Classic is the integration target. New Outlook and server-side conversion are not implemented equivalents. Automated tests of services do not replace live Word/Outlook qualification on both Office architectures.
