using System.Runtime.InteropServices;

namespace GBLauncher;

/// <summary>
/// Path validation and drive type detection.
/// </summary>
internal static class PathHelper
{
    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern uint GetDriveType(string lpRootPathName);

    private const uint DRIVE_REMOVABLE = 2;
    private const uint DRIVE_FIXED = 3;
    private const uint DRIVE_REMOTE = 4;
    private const uint DRIVE_CDROM = 5;
    private const uint DRIVE_RAMDISK = 6;

    public enum DriveType
    {
        Invalid,
        FixedDrive,
        RemovableDrive,
        CDRom
    }

    public static DriveType GetDriveTypeFromPath(string path)
    {
        if (string.IsNullOrEmpty(path) || path.Length < 3)
            return DriveType.Invalid;

        var rootPath = path.Substring(0, 3); // e.g., "C:\"
        var driveType = GetDriveType(rootPath);

        return driveType switch
        {
            DRIVE_FIXED or DRIVE_REMOTE or DRIVE_RAMDISK => DriveType.FixedDrive,
            DRIVE_CDROM => DriveType.CDRom,
            DRIVE_REMOVABLE => DriveType.RemovableDrive,
            _ => DriveType.Invalid
        };
    }

    public static string? ValidatePath(string path, string appPath)
    {
        if (string.IsNullOrEmpty(path))
            return null;

        var workingPath = path;

        // Convert root "\" to current drive root
        if (workingPath.StartsWith('\\') && !workingPath.StartsWith("\\\\"))
        {
            workingPath = appPath.Substring(0, 2) + workingPath;
        }

        // Convert relative path to absolute
        if (workingPath.Length > 1 && workingPath[1] != ':')
        {
            workingPath = Path.Combine(appPath, workingPath);
        }

        // Try to fully qualify the path
        try
        {
            workingPath = Path.GetFullPath(workingPath);
        }
        catch
        {
            return null;
        }

        // If specified path does not exist, try to validate on the current drive
        if (!Directory.Exists(workingPath))
        {
            var localPath = appPath.Substring(0, 2) + workingPath.Substring(2);
            if (Directory.Exists(localPath))
            {
                workingPath = localPath;
            }
        }

        // Verify path exists
        var driveType = GetDriveTypeFromPath(workingPath);
        if (driveType == DriveType.FixedDrive && !Directory.Exists(workingPath))
        {
            return null;
        }

        return workingPath;
    }
}
