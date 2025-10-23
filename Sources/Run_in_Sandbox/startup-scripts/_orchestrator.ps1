param(
    [string]$ScriptsPath = "C:\Run_in_Sandbox\startup-scripts",
    # Can include this switch when running from the .wsb file to indicate it's the first launch of the sandbox
    # Useful if re-running this script within the sandbox as a test, but don't want certain parts to run again
    [switch]$launchingSandbox
)

# ------ Check that we're running in the Windows Sandbox ------
# This script is intended to be run from within the Windows Sandbox. We'll do a rudamentary check for if the current user is named "WDAGUtilityAccount"
if ($env:USERNAME -ne "WDAGUtilityAccount") {
    Write-host "`n`nERROR: This script is intended to be run from WITHIN the Windows Sandbox.`nIt appears you are running this from outside the sandbox.`n" -ForegroundColor Red
    Write-host "`nPress Enter to exit." -ForegroundColor Yellow
    Read-Host
    exit
}

Write-Host "[Orchestrator] Scripts path: $ScriptsPath"

# 1) Run ordered startup scripts: 00-*, 01-* ... 99-*
$pattern = '^\d{2}-.+\.ps1$'
$items = Get-ChildItem -LiteralPath $ScriptsPath -Filter *.ps1 -File -ErrorAction SilentlyContinue |
         Where-Object { $_.Name -match $pattern } |
         Sort-Object Name

foreach ($i in $items) {
    Write-Host "[Orchestrator] Running: $($i.Name)"
    try {
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $i.FullName
        $rc = $LASTEXITCODE
        if ($rc -ne $null -and $rc -ne 0) {
            Write-Warning "[Orchestrator] Script $($i.Name) returned exit code $rc"
        }
    } catch {
        Write-Warning "[Orchestrator] Script $($i.Name) threw: $($_.Exception.Message)"
    }
}

# Restart Explorer so changes take effect
Write-Host "[Orchestrator] Restarting Explorer so changes take effect"
Get-Process explorer | Stop-Process -Force

# 2) Read and run the original command last
$origFile = Join-Path $ScriptsPath "OriginalCommand.txt"
if (Test-Path -LiteralPath $origFile) {
    $orig = Get-Content -LiteralPath $origFile -Raw
    Write-Host "[Orchestrator] Running original command..."
    # Run through cmd to support both cmd and PowerShell-style lines
    Start-process -Filepath "C:\Windows\SysWOW64\cmd.exe" -ArgumentList @('/c', '"' + $orig + '"') -WindowStyle Hidden
} else {
    Write-Warning "[Orchestrator] OriginalCommand.txt not found; nothing to run."
}
