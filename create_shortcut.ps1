$ws = New-Object -ComObject WScript.Shell
$desktop = [Environment]::GetFolderPath('Desktop')
$sc = $ws.CreateShortcut("$desktop\RE Underwriting Tool.lnk")
$sc.TargetPath = "C:\ai project folder\re-underwriting-tool-clone\start_server.bat"
$sc.WorkingDirectory = "C:\ai project folder\re-underwriting-tool-clone"
$sc.IconLocation = "C:\Windows\System32\shell32.dll,13"
$sc.Description = "Start RE Underwriting Tool Server"
$sc.Save()
Write-Output "Desktop shortcut created successfully!"
