# 构建便携版（双击即用，无需安装 Python）
# 用法：powershell -ExecutionPolicy Bypass -File packaging/build_portable.ps1

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

$py = if ($env:PYTHON) { $env:PYTHON } else { 'C:\Program Files\Python313\python.exe' }
$static = Join-Path $root 'src\excel2ppt_studio\web\static'

& $py -m PyInstaller `
    --noconfirm --clean `
    --name Excel2PPT-Studio `
    --onedir `
    --specpath build `
    --add-data "$static;excel2ppt_studio/web/static" `
    --collect-submodules uvicorn `
    --collect-submodules fastapi `
    --collect-submodules pydantic `
    --hidden-import uvicorn.logging `
    --hidden-import uvicorn.loops.auto `
    --hidden-import uvicorn.protocols.http.auto `
    --hidden-import uvicorn.protocols.websockets.auto `
    --hidden-import uvicorn.lifespan.on `
    packaging/web_launcher.py

if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller 失败，退出码 $LASTEXITCODE"
}

$dist = Join-Path $root 'dist\Excel2PPT-Studio'
$exe = Join-Path $dist 'Excel2PPT-Studio.exe'
if (-not (Test-Path $exe)) {
    throw "未生成入口：$exe"
}

$size = (Get-ChildItem $dist -Recurse -File | Measure-Object Length -Sum).Sum
Write-Host ""
Write-Host "产物目录：$dist"
Write-Host "入口：$exe"
Write-Host "体积：$([math]::Round($size / 1MB, 1)) MB"
