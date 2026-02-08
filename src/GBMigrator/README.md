# GBMigrator - GameBase Access to SQLite Migration Tool

## Overview
Command-line utility to migrate GameBase data from Microsoft Access (.mdb) to SQLite databases.

## Requirements
- .NET 8.0 Runtime (Windows)
- **Microsoft Access Database Engine** (2016 or later) or **Microsoft Jet 4.0 OLEDB Provider**

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

```bash
GBMigrator <access-db-path> <sqlite-db-path>
```

### Example
```bash
GBMigrator C:\GameBase\GameBase.mdb C:\Users\YourName\AppData\Roaming\GameBase\Data\GameBase.db
```

## What Gets Migrated

The tool migrates the following tables:

### Lookup Tables (13 tables)
- Publishers, Years, Programmers, Musicians, Languages
- Genres, PGenres (Parent Genres), Difficulty, Crackers, Artists
- Developers, Licenses, Rarities

### Main Tables
- **Games** - Complete game records with all fields
- **Extras** - Game extras/additional content

## Features

✅ **Preserves IDs** - All primary keys and foreign keys are maintained  
✅ **Progress reporting** - Real-time console output  
✅ **Schema detection** - Skips missing tables gracefully  
✅ **Overwrite protection** - Prompts before replacing existing SQLite database  
✅ **Null handling** - Correctly handles nullable fields from Access  

## Build from Source

```bash
cd src/GBMigrator
dotnet build
```

## Run from Source

```bash
dotnet run --project src/GBMigrator/GBMigrator.csproj -- "path/to/source.mdb" "path/to/target.db"
```

## Troubleshooting

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
