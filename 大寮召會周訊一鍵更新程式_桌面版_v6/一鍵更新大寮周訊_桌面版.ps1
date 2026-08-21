$ErrorActionPreference = 'Stop'
$version = '2026.08.15-v5'
$Host.UI.RawUI.WindowTitle = "大寮召會週訊一鍵更新 $version"
$appDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$desktop = [Environment]::GetFolderPath('Desktop')
$template = Join-Path $appDir '參考範本.docx'
$log = Join-Path $desktop '周訊更新執行記錄.txt'

try {
    @(
        "版本：$version"
        "執行時間：$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        "程式位置：$appDir"
        "輸出桌面：$desktop"
    ) | Set-Content -LiteralPath $log -Encoding UTF8

    Write-Host "版本：$version" -ForegroundColor DarkGray
    Write-Host "輸出位置：$desktop" -ForegroundColor Cyan

    if (-not (Test-Path -LiteralPath $template)) {
        throw '找不到「參考範本.docx」，請完整解壓縮整個 ZIP 後再執行。'
    }

    $pythonCandidates = Get-ChildItem -Path "$env:USERPROFILE\.cache\codex-runtimes" -Recurse -Filter python.exe -ErrorAction SilentlyContinue |
        Where-Object {
            $_.FullName -like '*codex-primary-runtime*dependencies*python*' -and
            $_.DirectoryName -like '*\dependencies\python'
        } |
        Sort-Object LastWriteTime -Descending

    if (-not $pythonCandidates) {
        throw '找不到 Codex 文件處理環境。請先安裝並開啟 Codex，再重新執行。'
    }

    $python = $pythonCandidates[0].FullName
    "Python：$python" | Add-Content -LiteralPath $log -Encoding UTF8
    Write-Host '正在重新讀取網站並建立周訊，請勿關閉視窗…' -ForegroundColor Cyan

    $output = & $python "$appDir\format_runner.py" --template "$template" --output-dir "$desktop" 2>&1
    $exitCode = $LASTEXITCODE
    $output | Tee-Object -FilePath $log -Append | ForEach-Object { Write-Host $_ }

    if ($exitCode -ne 0) {
        throw "更新程式傳回錯誤代碼 $exitCode。"
    }

    Write-Host "`n更新完成。Word 檔與摘要已放到：$desktop" -ForegroundColor Green
    Write-Host "執行記錄：$log" -ForegroundColor Green
}
catch {
    $message = $_.Exception.Message
    "失敗：$message" | Add-Content -LiteralPath $log -Encoding UTF8
    Write-Host "`n更新失敗：$message" -ForegroundColor Red
    Write-Host "請把這個檔案傳回協助檢查：$log" -ForegroundColor Yellow
}
finally {
    Write-Host ''
    Read-Host '按 Enter 關閉視窗'
}

