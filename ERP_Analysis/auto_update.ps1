# ──────────────────────────────────────────────────────────────
# 鴻騰整合儀表板 — 每週五 15:00 自動更新腳本
# ──────────────────────────────────────────────────────────────

$ROOT   = "C:\Users\USER\Honten"
$SCRIPT = "$ROOT\ERP_Analysis\dashboard.py"
$LOG    = "$ROOT\ERP_Analysis\auto_update.log"
$DATE   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

function Log($msg) {
    $line = "[$DATE] $msg"
    Write-Host $line
    Add-Content -Path $LOG -Value $line -Encoding UTF8
}

Log "===== 開始自動更新 ====="

# 1. 執行 Python 產生儀表板
Log "執行 dashboard.py..."
Set-Location "$ROOT\ERP_Analysis"
$result = python dashboard.py 2>&1
if ($LASTEXITCODE -ne 0) {
    Log "[ERROR] dashboard.py 執行失敗："
    Log $result
    exit 1
}
Log "dashboard.py 完成"

# 2. 找最新的 HTML 檔案
$html = Get-ChildItem "$ROOT\ERP_Analysis\output\鴻騰整合儀表板_*.html" |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $html) {
    Log "[ERROR] 找不到 HTML 輸出檔"
    exit 1
}
Log "產出檔案：$($html.Name)"

# 3. Git add / commit / push
Set-Location $ROOT

git add "ERP_Analysis/output/$($html.Name)"
git add "ERP_Analysis/*.py"
git add "ERP_Analysis/.gitignore"

$commitMsg = "auto: 每週儀表板更新 $(Get-Date -Format 'yyyy-MM-dd')"
git commit -m $commitMsg
if ($LASTEXITCODE -ne 0) {
    Log "[INFO] 無新變更，跳過 commit"
} else {
    Log "已 commit：$commitMsg"
}

git push origin main
if ($LASTEXITCODE -ne 0) {
    Log "[ERROR] git push 失敗"
    exit 1
}
Log "已 push 到 GitHub"
Log "===== 更新完成 ====="
