# PowerShell snippet to copy a locked Excel file on Windows
# Uses Win32 CreateFile with shared read/write/delete access

Add-Type @"
using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

public class LockedFileReader {
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    static extern SafeFileHandle CreateFile(
        string lpFileName,
        uint dwDesiredAccess,
        uint dwShareMode,
        IntPtr lpSecurityAttributes,
        uint dwCreationDisposition,
        uint dwFlagsAndAttributes,
        IntPtr hTemplateFile);

    public static byte[] ReadLockedFile(string path) {
        // GENERIC_READ = 0x80000000
        // FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE = 7
        // OPEN_EXISTING = 3
        SafeFileHandle handle = CreateFile(path, 0x80000000, 7, IntPtr.Zero, 3, 0, IntPtr.Zero);
        if (handle.IsInvalid) {
            throw new IOException("Cannot open file. Win32 Error: " + Marshal.GetLastWin32Error());
        }
        using (var fs = new FileStream(handle, FileAccess.Read)) {
            byte[] buffer = new byte[fs.Length];
            fs.Read(buffer, 0, buffer.Length);
            return buffer;
        }
    }
}
"@

# Usage:
# $sourcePath = (Resolve-Path "YourFile.xlsx").Path
# $bytes = [LockedFileReader]::ReadLockedFile($sourcePath)
# [System.IO.File]::WriteAllBytes("$PWD\YourFile_copy.xlsx", $bytes)
