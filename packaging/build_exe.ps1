# 打包 Windows 可执行程序（onedir，便于排查缺失依赖）
# 用法：powershell -ExecutionPolicy Bypass -File packaging/build_exe.ps1

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

$py = if ($env:PYTHON) { $env:PYTHON } else { 'C:\Program Files\Python313\python.exe' }
$static = Join-Path $root 'src\excel2ppt_studio\web\static'

& $py -m PyInstaller `
    --noconfirm --clean `
    --name e2p `
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
    packaging/e2p_launcher.py

if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller 失败，退出码 $LASTEXITCODE"
}

Write-Host ""
Write-Host "产物：$root\dist\e2p\e2p.exe"
Write-Host "试运行：& '$root\dist\e2p\e2p.exe' --help"
