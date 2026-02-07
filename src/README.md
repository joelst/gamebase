# GameBase.NET

This is a somewhat sloppy AI-assisted upgrade of GameBase 1.3 from VB6 to .NET 8. The orginal was a very thorough creation so just doing an upgrade is a big challenge.

## Archive Backend Configuration
- **Defaults**: Windows ➜ `SevenZipCli`; others ➜ `Managed7Zip`.
- **Env override**: `GEMUS_ARCHIVE_BACKEND` (aliases: `SevenZipCli`, `7zip`, `ManagedZip`, `Managed7Zip`, `sharpcompress`).
- **Gemus.ini** (`[Settings]`):
  ```ini
  [Settings]
  ArchiveBackend=Managed7Zip
  ```
- **UI**: GameBase ➜ **Options** dialog to set and persist `ArchiveBackend` to `Gemus.ini`.
