# GameBase.NET

GameBase.NET is a .NET 8 migration of the original GameBase 1.3 (VB6) retro‑gaming frontend and database utility. The repository includes the .NET source, the original VB6 source for reference, tooling for database migration, and supporting assets.

## Projects

### Core applications
- **GameBase** (WinForms): Main frontend
- **GBToolbox** (WinForms): Tools and utilities

### Utilities
- **GBMigrator**: Access ↔ SQLite migration and diff tooling
- **GBLauncher**: Launcher utility
- **GBRegRun**: Registration/runner utility
- **HVSC2GB64**: HVSC database update tool

### Libraries
- **GBFuncs**: Core helpers
- **GBSql**: Database helpers and query builder
- **Gemus**: Emulator scripting (GEMUS)

## Requirements

- **Windows** for full functionality (Access import/export requires Windows)
- **.NET 8 SDK**
- **Microsoft Access Database Engine 2016** for Access (.mdb) import/export
  - https://www.microsoft.com/en-us/download/details.aspx?id=54920

## Build

From the repository root:

```bash
# Build the main application
 dotnet build src/GameBase/GameBase.csproj

# Build migration tooling
 dotnet build src/GBMigrator/GBMigrator.csproj
```

## Tests

```bash
 dotnet test tests/GBSql.Tests/GBSql.Tests.csproj
 dotnet test tests/Gemus.Tests/Gemus.Tests.csproj
```

## Database Migration

### Import (Access → SQLite)

```bash
GBMigrator <access-db-path> <sqlite-db-path>
```

### Export (SQLite → Access)

```bash
GBMigrator export <sqlite-db-path> <access-db-path>
```

See [src/GBMigrator/README.md](src/GBMigrator/README.md) for full details.

## MSIX Installer

MSIX packaging is supported via a PowerShell script. It builds all apps in Release mode, stages the outputs, generates MSIX assets from the GameBase ICO, and produces a signed MSIX.

```powershell
./Setup/msix/Build-MSIX.ps1 `
  -packageName "GameBase.GameBaseNet" `
  -publisher "CN=NamedPipes Software, O=NamedPipes Consulting, C=US" `
  -publisherDisplayName "NamedPipes Software" `
  -displayName "GameBase.NET" `
  -version "1.0.0.0" `
  -appExecutable "GameBase.exe" `
  -iconPath "./Artwork/Program Icons/GameBase.ico" `
  -outputPath "Output" `
  -createTestCertificate
```

Prerequisites:
- Windows SDK (MakeAppx, SignTool)
- Microsoft Access Database Engine 2016 (for export tools)

## Repository Notes

- The **original VB6 source** is preserved under [Source](Source) for reference.
- Prebuilt binaries and installer outputs are intentionally excluded from source control.
- Artwork and historical assets remain under [Artwork](Artwork).

## License

See [License.txt](License.txt).
