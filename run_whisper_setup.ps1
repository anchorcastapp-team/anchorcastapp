# AnchorCast Whisper Setup Runner
# Disables the close button on this window so user cannot accidentally close it mid-install

Add-Type @"
using System;
using System.Runtime.InteropServices;
public class WindowHelper {
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]
    public static extern IntPtr GetSystemMenu(IntPtr hWnd, bool bRevert);
    [DllImport("user32.dll")]
    public static extern bool DeleteMenu(IntPtr hMenu, uint uPosition, uint uFlags);
    [DllImport("user32.dll")]
    public static extern bool DrawMenuBar(IntPtr hWnd);
    public const uint SC_CLOSE = 0xF060;
    public const uint MF_BYCOMMAND = 0x00000000;
}
"@

# Disable the close (X) button
$hwnd = [WindowHelper]::GetConsoleWindow()
$hmenu = [WindowHelper]::GetSystemMenu($hwnd, $false)
[WindowHelper]::DeleteMenu($hmenu, [WindowHelper]::SC_CLOSE, [WindowHelper]::MF_BYCOMMAND)
[WindowHelper]::DrawMenuBar($hwnd)

# Set window title
$host.UI.RawUI.WindowTitle = "AnchorCast - Whisper AI Setup"

# Run the bat file
$batPath = $args[0]
if (-not $batPath) {
    Write-Host "ERROR: No bat path provided"
    exit 1
}

$env:ANCHORCAST_NONINTERACTIVE = "0"
$process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$batPath`"" -Wait -NoNewWindow -PassThru
exit $process.ExitCode
