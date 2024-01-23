# Add-Type to access load .NET Framework Classes
Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;

    public class Clipboard
    {
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool CloseClipboard();

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr GetClipboardData(uint uFormat);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool IsClipboardFormatAvailable(uint format);

        [DllImport("Kernel32.dll", SetLastError = true)]
        public static extern IntPtr GlobalLock(IntPtr hMem);

        [DllImport("Kernel32.dll", SetLastError = true)]
        public static extern bool GlobalUnlock(IntPtr hMem);
    }
"@

# Load the assembly required for sending keystrokes
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

# Define clipboard format
$CF_UNICODETEXT = 13

# Check if clipboard format is available
if ([Clipboard]::IsClipboardFormatAvailable($CF_UNICODETEXT) -eq $false) {
    Write-Host "No text data is available in the clipboard"
    exit
}

# Open clipboard
if ([Clipboard]::OpenClipboard([IntPtr]::Zero) -eq $false) {
    Write-Host "Unable to open clipboard"
    exit
}

# Get clipboard data
$handle = [Clipboard]::GetClipboardData($CF_UNICODETEXT)
if ($handle -eq [IntPtr]::Zero) {
    Write-Host "Unable to get clipboard data"
    [Clipboard]::CloseClipboard()
    exit
}

# Lock clipboard data
$pointer = [Clipboard]::GlobalLock($handle)
if ($pointer -eq [IntPtr]::Zero) {
    Write-Host "Unable to lock clipboard data"
    [Clipboard]::CloseClipboard()
    exit
}

# Get clipboard text
$clipboardText = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($pointer)

# Unlock clipboard data
[Clipboard]::GlobalUnlock($handle)

# Close clipboard
[Clipboard]::CloseClipboard()

# Wait for 3 seconds to give the user time to focus the application
Start-Sleep -Seconds 3

# Send clipboard text as keystrokes to the currently focused application
[System.Windows.Forms.SendKeys]::SendWait($clipboardText)

if ($?) {
    Write-Host "SendKeys successful"
}
else {
    # Wait for 10 seconds to give the user time to notice the error
    Write-Host "SendKeys failed"
    Start-Sleep -Seconds 10    
}