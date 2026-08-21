$ErrorActionPreference = 'Stop'
$version = '2026.08.15-v6'
$Host.UI.RawUI.WindowTitle = "Weekly Bulletin Updater $version"
$appDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$desktop = [Environment]::GetFolderPath('Desktop')
$template = Join-Path $appDir 'reference.docx'
$log = Join-Path $desktop 'weekly_bulletin_update_log.txt'

try {
    @(
        "Version: $version"
        "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        "Program: $appDir"
        "Output: $desktop"
    ) | Set-Content -LiteralPath $log -Encoding UTF8

    Write-Host "Version: $version" -ForegroundColor DarkGray
    Write-Host "Output folder: $desktop" -ForegroundColor Cyan
    if (-not (Test-Path -LiteralPath $template)) {
        throw 'reference.docx is missing. Please extract the complete ZIP before running.'
    }

    $pythonCandidates = Get-ChildItem -Path "$env:USERPROFILE\.cache\codex-runtimes" -Recurse -Filter python.exe -ErrorAction SilentlyContinue |
        Where-Object {
            $_.FullName -like '*codex-primary-runtime*dependencies*python*' -and
            $_.DirectoryName -like '*\dependencies\python'
        } |
        Sort-Object LastWriteTime -Descending
    if (-not $pythonCandidates) {
        throw 'Codex document runtime was not found. Open Codex once on this computer, then try again.'
    }

    $python = $pythonCandidates[0].FullName
    "Python: $python" | Add-Content -LiteralPath $log -Encoding UTF8
    Write-Host 'Downloading the current bulletin and creating Word file...' -ForegroundColor Cyan
    $result = & $python "$appDir\format_runner.py" --template "$template" --output-dir "$desktop" 2>&1
    $exitCode = $LASTEXITCODE
    $result | Tee-Object -FilePath $log -Append | ForEach-Object { Write-Host $_ }
    if ($exitCode -ne 0) {
        throw "Updater exited with code $exitCode."
    }

    Write-Host ''
    Write-Host "Completed. Check the Word file on: $desktop" -ForegroundColor Green
    Write-Host "Log: $log" -ForegroundColor Green
}
catch {
    $message = $_.Exception.Message
    "ERROR: $message" | Add-Content -LiteralPath $log -Encoding UTF8
    Write-Host ''
    Write-Host "Update failed: $message" -ForegroundColor Red
    Write-Host "Please send this log file for diagnosis: $log" -ForegroundColor Yellow
}
finally {
    Write-Host ''
    Read-Host 'Press Enter to close'
}
