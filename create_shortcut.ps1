$WshShell = New-Object -comObject WScript.Shell
$DesktopPath = [Environment]::GetFolderPath("Desktop")
$Shortcut = $WshShell.CreateShortcut("$DesktopPath\Compilador.lnk")
$Shortcut.TargetPath = "$PWD\start_hidden.vbs"
$Shortcut.IconLocation = "notepad.exe,0"
$Shortcut.Description = "Iniciar Compilador Excel"
$Shortcut.WorkingDirectory = "$PWD"
$Shortcut.Save()
Write-Host "Shortcut created on Desktop: $HOME\Desktop\Compilador.lnk"
