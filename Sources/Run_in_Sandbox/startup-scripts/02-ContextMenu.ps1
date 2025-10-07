# Change context menu to old style
reg.exe add "HKCU\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32" /f /ve

# Show file extensions
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "HideFileExt" /t REG_DWORD /d 0 /f

# Show hidden files
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "Hidden" /t REG_DWORD /d 1 /f



# ---- Add 'Open PowerShell Here' and 'Open CMD Here' to context menu -------
$powershellPath = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
$cmdPath = "C:\Windows\SysWOW64\cmd.exe"
Write-Host "`nAdding 'Open PowerShell/CMD Here' context menu options"
reg add "HKEY_CLASSES_ROOT\Directory\Background\shell\MyPowerShell" /ve /d "Open PowerShell Here" /f
reg add "HKEY_CLASSES_ROOT\Directory\Background\shell\MyPowerShell" /v "Icon" /t REG_SZ /d "$powershellPath,0" /f
reg add "HKEY_CLASSES_ROOT\Directory\Background\shell\MyPowerShell\command" /ve /d "powershell.exe -noexit -command Set-Location -literalPath '%V'" /f
reg add "HKEY_CLASSES_ROOT\Directory\Background\shell\Mycmd" /ve /d "Open CMD Here" /f
reg add "HKEY_CLASSES_ROOT\Directory\Background\shell\Mycmd" /v "Icon" /t REG_SZ /d "$cmdPath,0" /f
reg add "HKEY_CLASSES_ROOT\Directory\Background\shell\Mycmd\command" /ve /d "cmd.exe /s /k cd /d `"\`"%V`"\`"" /f

# ---- Add File Types to Context Menu > New ----
# ShellNew Text Document - .txt
Write-host "`nAdding txt document new file option"
reg add "HKEY_CLASSES_ROOT\txtfile" /ve /d "Text Document" /f
reg add "HKEY_CLASSES_ROOT\.txt\ShellNew" /f
# Use --% to not have powershell parse the arguments, otherwise it won't pass the empty string for the /d parameter
reg --% add "HKEY_CLASSES_ROOT\.txt\ShellNew" /v "NullFile" /t REG_SZ /d "" /f
reg add "HKEY_CLASSES_ROOT\.txt\ShellNew" /v "ItemName" /t REG_SZ /d "New Text Document" /f

# ShellNew PowerShell Script - .ps1 -- Also happens to make .ps1 scripts clickable to run because of the association with "ps1file"
Write-host "`nAdding PowerShell new file option"
reg add "HKEY_CLASSES_ROOT\.ps1" /ve /d "ps1file" /f
reg add "HKEY_CLASSES_ROOT\ps1file" /ve /d "PowerShell Script" /f
reg add "HKEY_CLASSES_ROOT\ps1file\DefaultIcon" /ve /d "%SystemRoot%\System32\imageres.dll,-5372" /f
reg add "HKEY_CLASSES_ROOT\.ps1\ShellNew" /ve /d "ps1file" /f
reg add "HKEY_CLASSES_ROOT\.ps1\ShellNew" /f
reg --% add "HKEY_CLASSES_ROOT\.ps1\ShellNew" /v "NullFile" /t REG_SZ /d "" /f
reg add "HKEY_CLASSES_ROOT\.ps1\ShellNew" /v "ItemName" /t REG_SZ /d "script" /f