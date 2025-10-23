# Fix for slow MSI package install. See: https://github.com/microsoft/Windows-Sandbox/issues/68#issuecomment-2754867968
reg add "HKLM\SYSTEM\CurrentControlSet\Control\CI\Policy" /v "VerifiedAndReputablePolicyState" /t REG_DWORD /d 0 /f
CiTool.exe --refresh --json | Out-Null # Refreshes policy. Use json output param or else it will prompt for confirmation, even with Out-Null