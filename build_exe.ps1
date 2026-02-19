# FileMonitor exe 빌드 스크립트 (PyInstaller)
# 사용: .\build_exe.ps1

$ErrorActionPreference = "Stop"
$ProjectRoot = $PSScriptRoot
$SpecName = "file_monitor.spec"

Set-Location $ProjectRoot

# 가상환경 활성화는 사용자가 수행 (선택)
# & .\.venv\Scripts\Activate.ps1

Write-Host "PyInstaller로 exe 빌드 중..." -ForegroundColor Cyan
Write-Host "  프로젝트: $ProjectRoot" -ForegroundColor Gray
Write-Host "  spec: $SpecName" -ForegroundColor Gray

& python -m PyInstaller --noconfirm $SpecName

if ($LASTEXITCODE -ne 0) {
    Write-Host "빌드 실패 (exit $LASTEXITCODE)" -ForegroundColor Red
    exit $LASTEXITCODE
}

$OutExe = Join-Path $ProjectRoot "dist\FileMonitor.exe"
if (Test-Path $OutExe) {
    Write-Host "`n빌드 완료: $OutExe" -ForegroundColor Green
} else {
    Write-Host "`n경고: dist\FileMonitor.exe 를 찾을 수 없습니다." -ForegroundColor Yellow
}
