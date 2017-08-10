#Lord Hagen / olehag04@nfk.no

#Continue if errors occurs.
$ErrorActionPreference = "silentlycontinue"

#This is required for system box message. That last line of code.
Add-Type -AssemblyName PresentationFramework

#If the proscess(es) are closed, remove the file.
if ((Get-Process lync, outlook) -eq $null) 
{
    Remove-Item $env\Users\$env:USERNAME\AppData\Local\Microsoft\outlook\*.ost
}

#If not, try to exit them. Then delete the file.
else 
{
# Find Skype process.
    $lync = Get-Process lync
    if ($lync) {
# Try to exit.
    $lync.CloseMainWindow()
# Wait 5 seconds.
    Start-Sleep 5
# If still running, kill the process.
    if (!$lync.HasExited) 
    {$lync | Stop-Process -Force}
}
Remove-Variable lync

# Find Outlook process.
    $outlook = Get-Process outlook
    if ($outlook) {
# Try to exit.
    $outlook.CloseMainWindow()
# Wait 5 seconds.
    Start-Sleep 5
# If still running, kill the process.
    if (!$outlook.HasExited) 
    {$outlook | Stop-Process -Force}
}
Remove-Variable outlook

Remove-Item $env\Users\$env:USERNAME\AppData\Local\Microsoft\outlook\*.ost
}

[System.Windows.MessageBox]::Show('Start Outlook!')
