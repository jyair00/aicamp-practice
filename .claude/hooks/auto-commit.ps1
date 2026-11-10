# Auto-commit hook: runs when Claude finishes a turn (Stop hook).
# Commits every change in this folder that .gitignore does not exclude.
# It never pushes anything to a remote.
# Keep 'Continue': git writes harmless warnings (LF/CRLF) to stderr, and
# 'Stop' would turn them into errors and skip the commit.
$ErrorActionPreference = 'Continue'
try {
    $root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    $git = Join-Path $env:ProgramFiles 'Git\cmd\git.exe'
    if (-not (Test-Path -LiteralPath $git)) { $git = 'git' }
    Set-Location -LiteralPath $root

    & $git -c core.safecrlf=false add -A *> $null
    $files = @(& $git diff --cached --name-only)
    if ($files.Count -eq 0) { exit 0 }

    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $msgFile = Join-Path $root '.git\AUTO_COMMIT_MSG'
    $body = "auto-commit: $stamp ($($files.Count) files)`n`n" + (($files | Select-Object -First 30) -join "`n") + "`n"
    [IO.File]::WriteAllText($msgFile, $body, (New-Object Text.UTF8Encoding($false)))
    & $git commit -q -F $msgFile *> $null
    $ok = ($LASTEXITCODE -eq 0)
    Remove-Item -LiteralPath $msgFile -ErrorAction SilentlyContinue
    if (-not $ok) { Write-Output '{"systemMessage":"Auto-commit failed - check git status"}'; exit 0 }

    $hash = (& $git rev-parse --short HEAD)
    Write-Output ('{"systemMessage":"Auto-commit ' + $hash + ': ' + $files.Count + ' files"}')
} catch {
    exit 0
}
