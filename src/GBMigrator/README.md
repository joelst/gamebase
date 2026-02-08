# GBMigrator - GameBase Import/Export Tool

## Overview
Command-line utility to import GameBase data from Microsoft Access (.mdb) to SQLite and export back to Access for backward compatibility.

## Requirements
- **.NET 8.0 Runtime** (Windows only)
- **Windows OS** - Required for OleDb Access database provider
- **Microsoft Access Database Engine** (2016 or later) or **Microsoft Jet 4.0 OLEDB Provider**
- **ADOX** (installed with the Access Database Engine) for Access export

⚠️ **Platform Limitation**: This tool can only run on Windows because the OleDb provider for Microsoft Access databases is Windows-only.

### For Non-Windows Users
If you need to use GameBase on Linux or macOS:
1. Run GBMigrator on a Windows machine first to convert your Access database to SQLite
2. Copy the resulting `.db` file to your target system
3. Use the GameBase application with the SQLite database (cross-platform compatible)

### Installing Access Database Engine
If you get an error like:
```
The 'Microsoft.ACE.OLEDB.12.0' provider is not registered on the local machine.
```

Download and install the **Microsoft Access Database Engine 2016 Redistributable**:
- **64-bit**: https://www.microsoft.com/en-us/download/details.aspx?id=54920
- **32-bit**: https://www.microsoft.com/en-us/download/details.aspx?id=13255

⚠️ **Important**: Match the architecture (32-bit vs 64-bit) with your .NET runtime.

## Usage

### Access to SQLite Migration

```bash
GBMigrator <access-db-path> <sqlite-db-path>
```

Example:

```bash
GBMigrator C:\GameBase\GameBase.mdb C:\Users\YourName\AppData\Roaming\GameBase\Data\GameBase.db
```

### SQLite to Access Export

```bash
GBMigrator export <sqlite-db-path> <access-db-path>
```

Example:

```bash
GBMigrator export C:\Users\YourName\AppData\Roaming\GameBase\Data\GameBase.db C:\GameBase\GameBase.mdb
```

### Diff Utility (SQLite)

```bash
GBMigrator diff <current-db> <backup-db> [output-db] [--letters <letters>]
```

Examples:

```bash
GBMigrator diff C:\GameBase\Current.db C:\GameBase\Backup.db
GBMigrator diff Current.db Backup.db update.db --letters A,B,C,0-9
```

Letters:
- Use `A-Z` to restrict by starting letter.
- Use `0-9` to include non-letter names.
- If omitted, all letters and `0-9` are included.

## What Gets Migrated

The tool migrates the following tables:

### Lookup Tables (13 tables)
- Publishers, Years, Programmers, Musicians, Languages
- Genres, PGenres (Parent Genres), Difficulty, Crackers, Artists
- Developers, Licenses, Rarities

### Main Tables
- **Games** - Complete game records with all fields
- **Extras** - Game extras/additional content

## What Gets Exported

The export uses the same table coverage as import:

### Lookup Tables (13 tables)
- Publishers, Years, Programmers, Musicians, Languages
- Genres, PGenres (Parent Genres), Difficulty, Crackers, Artists
- Developers, Licenses, Rarities

### Main Tables
- **Games** - Includes `ScreenshotFilename` when present
- **Extras** - Uses `EX_Type` (matches SQLite schema)

## Features

✅ **Preserves IDs** - All primary keys and foreign keys are maintained
✅ **Progress reporting** - Real-time console output
✅ **Schema detection** - Skips missing tables gracefully
✅ **Overwrite protection** - Prompts before replacing existing SQLite database
✅ **Null handling** - Correctly handles nullable fields from Access
✅ **Backward compatibility** - Export produces Access `.mdb` for VB6 GameBase

## Build from Source

```bash
cd src/GBMigrator
dotnet build
```

## Run from Source

```bash
dotnet run --project src/GBMigrator/GBMigrator.csproj -- "path/to/source.mdb" "path/to/target.db"
```

Run the diff tool from source:

```bash
dotnet run --project src/GBMigrator/GBMigrator.csproj -- diff "path/to/current.db" "path/to/backup.db" "path/to/update.db"
```

Run export from source:

```bash
dotnet run --project src/GBMigrator/GBMigrator.csproj -- export "path/to/source.db" "path/to/target.mdb"
```

## Troubleshooting

### "GBMigrator requires Windows to access Microsoft Access databases"
This tool can only run on Windows. The OleDb provider for Access files is not available on other platforms.

**Solution**: Run the migration on a Windows machine, then copy the resulting `.db` file to your target system.

### "Microsoft Access Database Engine not found"
The required OleDb provider is not installed.

**Solution**: Install the Microsoft Access Database Engine 2016 Redistributable:
- Download: https://www.microsoft.com/en-us/download/details.aspx?id=54920
- Install the version matching your system architecture (x64 or x86)

### "Cannot open a database created with a previous version"
The .mdb file format is incompatible. Try:
1. Installing Access Database Engine 2016
2. Opening and repairing the database in Microsoft Access

### "Provider is not registered"
Install the Microsoft Access Database Engine (see Requirements section above).

### Migration is slow
Large databases (>10,000 games) may take several minutes. Progress is reported every 500 records.

## Technical Notes

- Uses `System.Data.OleDb` for reading Access databases
- Uses `Microsoft.Data.Sqlite` for writing SQLite databases
- Creates SQLite database with proper schema automatically
- Migrates lookup tables first to satisfy foreign key constraints
- Batch commits every record for reliability
