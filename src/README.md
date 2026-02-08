# GameBase.NET

This is the .NET 8 migration of the original GameBase 1.3 (VB6). For full project details, build instructions, and tooling notes, see the repository root README.

## Archive Backend Configuration
- **Defaults**: Windows ➜ `SevenZipCli`; others ➜ `Managed7Zip`.
- **Env override**: `GEMUS_ARCHIVE_BACKEND` (aliases: `SevenZipCli`, `7zip`, `ManagedZip`, `Managed7Zip`, `sharpcompress`).
- **7z CLI override**: `GEMUS_SEVENZIP_PATH` (full path to `7z.exe`).
- **Gemus.ini** (`[Settings]`):
  ```ini
  [Settings]
  ArchiveBackend=Managed7Zip
  ```
- **UI**: GameBase ➜ **Options** dialog to set and persist `ArchiveBackend` to `Gemus.ini`.
