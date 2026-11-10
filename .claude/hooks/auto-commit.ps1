# Auto-commit hook: runs when Claude finishes a turn (Stop hook).
# Commits every change in this folder that .gitignore does not exclude.
# It never pushes anything to a remote.
$ErrorActionPreference = 'Stop'
try {
    $root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    $git = Join-Path $env:ProgramFiles 'Git\cmd\git.exe'
    if (-not (Test-Path -LiteralPath $git)) { $git = 'git' }
    Set-Location -LiteralPath $root

    & $git add -A 2>$null
    $files = @(& $git diff --cached --name-only)
    if ($files.Count -eq 0) { exit 0 }

    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $msgFile = Join-Path $root '.git\AUTO_COMMIT_MSG'
    $body = "auto-commit: $stamp ($($files.Count) files)`n`n" + (($files | Select-Object -First 30) -join "`n") + "`n"
    [IO.File]::WriteAllText($msgFile, $body, (New-Object Text.UTF8Encoding($false)))
    & $git commit -q -F $msgFile 2>$null
    Remove-Item -LiteralPath $msgFile -ErrorAction SilentlyContinue

    $hash = (& $git rev-parse --short HEAD)
    Write-Output ('{"systemMessage":"Auto-commit ' + $hash + ': ' + $files.Count + ' files"}')
} catch {
    exit 0
}
