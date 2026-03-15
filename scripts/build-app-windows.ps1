# ---------------------------------------------------------------------------
# build-app-windows.ps1 — Build "Claude Word Assistant" for Windows
#
# Output: dist\ClaudeWordAssistant\ClaudeWordAssistant.exe
#
# Requirements (dev machine only):
#   - Python 3.8+ on PATH
#   - Internet access (to install build dependencies)
#   - Git for Windows recommended (provides openssl.exe for end users)
#
# Usage:
#   .\scripts\build-app-windows.ps1
#   .\scripts\build-app-windows.ps1 -Clean    # remove previous build first
# ---------------------------------------------------------------------------
param([switch]$Clean)

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Root      = Split-Path -Parent $ScriptDir
$BuildVenv = Join-Path $Root ".build-venv"
$DistDir   = Join-Path $Root "dist"
$AppName   = "Claude Word Assistant"
$ExeName   = "ClaudeWordAssistant"

Write-Host ""
Write-Host "=== Building $AppName ===" -ForegroundColor Cyan
Write-Host ""

# ---------------------------------------------------------------------------
# Clean
# ---------------------------------------------------------------------------
if ($Clean) {
    Write-Host "--- Cleaning previous build ---"
    @($BuildVenv, "$Root\.build-work", "$Root\.build-spec", $DistDir) | ForEach-Object {
        if (Test-Path $_) { Remove-Item -Recurse -Force $_ }
    }
    Write-Host "Clean done." -ForegroundColor Green
    Write-Host ""
}

# ---------------------------------------------------------------------------
# 1. Build venv
# ---------------------------------------------------------------------------
Write-Host "--- Setting up build environment ---"
python -m venv $BuildVenv
& "$BuildVenv\Scripts\pip.exe" install --upgrade pip --quiet
& "$BuildVenv\Scripts\pip.exe" install `
    "fastapi==0.115.0" `
    "uvicorn[standard]==0.34.0" `
    "python-dotenv==1.0.1" `
    "sse-starlette==2.1.3" `
    "pystray" `
    "pillow" `
    "pyinstaller" `
    --quiet
Write-Host "Build dependencies installed." -ForegroundColor Green
Write-Host ""

# ---------------------------------------------------------------------------
# 2. Generate .ico icon from the existing PNG
# ---------------------------------------------------------------------------
Write-Host "--- Generating .ico icon ---"
$IconSrc = Join-Path $Root "addin\assets\icon-80.png"
$IconDst = Join-Path $Root ".build-icon.ico"

& "$BuildVenv\Scripts\python.exe" -c @"
from PIL import Image
img = Image.open(r'$IconSrc').convert('RGBA')
img.save(r'$IconDst', format='ICO', sizes=[(16,16),(32,32),(48,48),(64,64),(128,128),(256,256)])
print('Icon saved.')
"@
Write-Host "Icon generated." -ForegroundColor Green
Write-Host ""

# ---------------------------------------------------------------------------
# 3. PyInstaller
# ---------------------------------------------------------------------------
Write-Host "--- Building executable with PyInstaller ---"
Push-Location $Root

& "$BuildVenv\Scripts\pyinstaller.exe" `
    --noconfirm `
    --onedir `
    --windowed `
    --name $ExeName `
    --icon $IconDst `
    --distpath $DistDir `
    --workpath "$Root\.build-work" `
    --specpath "$Root\.build-spec" `
    --add-data "addin;addin" `
    --add-data "server\main.py;." `
    --add-data "server\document.py;." `
    --add-data "server\session.py;." `
    --hidden-import "uvicorn.logging" `
    --hidden-import "uvicorn.loops" `
    --hidden-import "uvicorn.loops.auto" `
    --hidden-import "uvicorn.protocols" `
    --hidden-import "uvicorn.protocols.http" `
    --hidden-import "uvicorn.protocols.http.auto" `
    --hidden-import "uvicorn.protocols.websockets" `
    --hidden-import "uvicorn.protocols.websockets.auto" `
    --hidden-import "uvicorn.lifespan" `
    --hidden-import "uvicorn.lifespan.on" `
    --hidden-import "fastapi" `
    --hidden-import "sse_starlette" `
    --hidden-import "pystray._win32" `
    --hidden-import "PIL._imaging" `
    "server\app_launcher_windows.py"

Pop-Location
Write-Host "Executable built." -ForegroundColor Green
Write-Host ""

# ---------------------------------------------------------------------------
# 4. Bundle openssl.exe if available (zero-dependency cert generation)
# ---------------------------------------------------------------------------
Write-Host "--- Bundling openssl.exe ---"
$OpenSSLPaths = @(
    "C:\Program Files\Git\usr\bin\openssl.exe",
    "C:\Program Files (x86)\Git\usr\bin\openssl.exe"
)
$OpenSSL = $OpenSSLPaths | Where-Object { Test-Path $_ } | Select-Object -First 1

if ($OpenSSL) {
    # PyInstaller 6+ puts bundled files in _internal\, earlier versions in the root
    $InternalDir = Join-Path $DistDir "$ExeName\_internal"
    $FallbackDir = Join-Path $DistDir $ExeName
    $DestDir = if (Test-Path $InternalDir) { $InternalDir } else { $FallbackDir }
    Copy-Item $OpenSSL $DestDir -Force
    Write-Host "Bundled openssl.exe from: $OpenSSL" -ForegroundColor Green
} else {
    Write-Host "openssl.exe not found — users need Git for Windows installed." -ForegroundColor Yellow
}
Write-Host ""

# ---------------------------------------------------------------------------
# 5. Cleanup
# ---------------------------------------------------------------------------
@($BuildVenv, "$Root\.build-work", "$Root\.build-spec", $IconDst) | ForEach-Object {
    if (Test-Path $_) { Remove-Item -Recurse -Force $_ }
}

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
$ExePath = Join-Path $DistDir "$ExeName\$ExeName.exe"
$SizeMB  = [math]::Round((Get-Item $ExePath).Length / 1MB, 1)

Write-Host "=== Build complete ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "Executable : $ExePath"
Write-Host "Size       : $SizeMB MB"
Write-Host ""
Write-Host "To distribute:"
Write-Host "  cd '$DistDir'"
Write-Host "  Compress-Archive -Path '$ExeName' -DestinationPath '$AppName.zip'"
Write-Host "  Upload '$AppName.zip' to a GitHub Release"
Write-Host ""
Write-Host "Users: download the zip, extract, double-click $ExeName.exe"
Write-Host ""
