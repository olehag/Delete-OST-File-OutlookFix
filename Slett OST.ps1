#Lord Hagen / olehag04@nfk.no

$ErrorActionPreference = "silentlycontinue"
Add-Type -AssemblyName PresentationFramework

if ((Get-Process lync, outlook) -eq $null) 
{
    Remove-Item $env\Users\$env:USERNAME\AppData\Local\Microsoft\outlook\*.ost #-WhatIf
}

else 
{
# Finn skype Prosess.
    $lync = Get-Process lync
    if ($lync) {
# Prøv å avslutte.
    $lync.CloseMainWindow()
# Vent 5 sec.
    Start-Sleep 5
# Drep Prosessen.
    if (!$lync.HasExited) 
    {$lync | Stop-Process -Force}
}
Remove-Variable lync

# Finn skype Prosess.
    $outlook = Get-Process outlook
    if ($outlook) {
# Prøv å avslutte.
    $outlook.CloseMainWindow()
# Vent 5 sec.
    Start-Sleep 5
# Drep Prosessen
    if (!$outlook.HasExited) 
    {$outlook | Stop-Process -Force}
}
Remove-Variable outlook

Remove-Item $env\Users\$env:USERNAME\AppData\Local\Microsoft\outlook\*.ost #-WhatIf
}
[System.Windows.MessageBox]::Show('Start Outlook på nytt!')