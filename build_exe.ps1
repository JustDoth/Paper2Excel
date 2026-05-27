param(
    [string]$Python = "D:\ProgramFiles\Anaconda3\envs\paper2excel\python.exe",
    [string]$Version = "v0.1.0"
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ReleaseRoot = Join-Path $ProjectRoot "release"
$ReleaseApp = Join-Path $ReleaseRoot "Paper2Excel"
$ReleaseZip = Join-Path $ReleaseRoot ("Paper2Excel-{0}-windows.zip" -f $Version)
$EnvRoot = Split-Path -Parent $Python
$EnvLibraryBin = Join-Path $EnvRoot "Library\bin"
$EnvScripts = Join-Path $EnvRoot "Scripts"
$IconPath = Join-Path $ProjectRoot "assets\Paper2Excel.ico"

if (!(Test-Path $Python)) {
    throw "Python not found: $Python"
}
if (!(Test-Path $IconPath)) {
    throw "Icon not found: $IconPath"
}

Set-Location $ProjectRoot
$env:PYTHONNOUSERSITE = "1"
$env:PATH = "$EnvRoot;$EnvLibraryBin;$EnvScripts;$env:PATH"

& $Python -c "import requests, urllib3, certifi, charset_normalizer, idna; print('requests preflight OK:', requests.__version__)"
& $Python -c "import ssl; print('ssl preflight OK:', ssl.OPENSSL_VERSION)"
& $Python -m pytest -q

if (Test-Path (Join-Path $ProjectRoot "build")) {
    Remove-Item -LiteralPath (Join-Path $ProjectRoot "build") -Recurse -Force
}
if (Test-Path (Join-Path $ProjectRoot "dist")) {
    Remove-Item -LiteralPath (Join-Path $ProjectRoot "dist") -Recurse -Force
}
if (Test-Path $ReleaseApp) {
    Remove-Item -LiteralPath $ReleaseApp -Recurse -Force
}
if (Test-Path $ReleaseZip) {
    Remove-Item -LiteralPath $ReleaseZip -Force
}

$AssetData = "assets\Paper2Excel.ico;assets"

& $Python -m PyInstaller `
    --noconfirm `
    --clean `
    --onedir `
    --windowed `
    --name Paper2Excel `
    --icon $IconPath `
    --add-data $AssetData `
    --hidden-import requests `
    --hidden-import fitz `
    --collect-all pymupdf `
    --exclude-module pandas `
    --exclude-module numpy `
    --exclude-module pytest `
    --exclude-module pygments `
    --exclude-module matplotlib `
    --exclude-module scipy `
    main.py

New-Item -ItemType Directory -Force -Path $ReleaseRoot | Out-Null
Move-Item -LiteralPath (Join-Path $ProjectRoot "dist\Paper2Excel") -Destination $ReleaseApp
Copy-Item -LiteralPath (Join-Path $EnvLibraryBin "libssl-3-x64.dll") -Destination (Join-Path $ReleaseApp "_internal\libssl-3-x64.dll") -Force
Copy-Item -LiteralPath (Join-Path $EnvLibraryBin "libcrypto-3-x64.dll") -Destination (Join-Path $ReleaseApp "_internal\libcrypto-3-x64.dll") -Force
Copy-Item -LiteralPath (Join-Path $ProjectRoot "README.md") -Destination (Join-Path $ReleaseApp "README.md") -Force
Copy-Item -LiteralPath (Join-Path $ProjectRoot "README.zh-CN.md") -Destination (Join-Path $ReleaseApp "README.zh-CN.md") -Force
Copy-Item -LiteralPath (Join-Path $ProjectRoot "LICENSE") -Destination (Join-Path $ReleaseApp "LICENSE") -Force
Copy-Item -LiteralPath (Join-Path $ProjectRoot "config.example.json") -Destination (Join-Path $ReleaseApp "config.example.json") -Force
Copy-Item -LiteralPath (Join-Path $ProjectRoot "templates") -Destination (Join-Path $ReleaseApp "templates") -Recurse -Force
Copy-Item -LiteralPath (Join-Path $ProjectRoot "assets") -Destination (Join-Path $ReleaseApp "assets") -Recurse -Force
Copy-Item -LiteralPath (Join-Path $ProjectRoot "docs") -Destination (Join-Path $ReleaseApp "docs") -Recurse -Force

New-Item -ItemType Directory -Force -Path (Join-Path $ReleaseApp "outputs") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $ReleaseApp "logs") | Out-Null

$SelfTestPath = Join-Path $ReleaseApp "self_test.json"
& (Join-Path $ReleaseApp "Paper2Excel.exe") --self-test $SelfTestPath
if ($LASTEXITCODE -ne 0) {
    throw "Packaged EXE self-test failed. See $SelfTestPath"
}

$ForbiddenUserConfig = Join-Path $ReleaseApp "user_config.json"
if (Test-Path $ForbiddenUserConfig) {
    throw "Release contains user_config.json. Remove local secrets before publishing."
}

$SecretFiles = Get-ChildItem -LiteralPath $ReleaseApp -Recurse -File -Include *.json,*.txt,*.md,*.env
foreach ($File in $SecretFiles) {
    $Text = Get-Content -LiteralPath $File.FullName -Raw -ErrorAction SilentlyContinue
    if ($File.Name -ne "config.example.json" -and $Text -match '"api_key"\s*:\s*"[^"]+"') {
        throw "Possible API key found in release file: $($File.FullName)"
    }
    if ($Text -match 'sk-[A-Za-z0-9_\-]{20,}') {
        throw "Possible OpenAI-style key found in release file: $($File.FullName)"
    }
    if ($Text -match 'Bearer\s+[A-Za-z0-9_\-\.]{20,}') {
        throw "Possible bearer token found in release file: $($File.FullName)"
    }
}

Compress-Archive -LiteralPath $ReleaseApp -DestinationPath $ReleaseZip -Force

Write-Host "Built portable app:" $ReleaseApp
Write-Host "Built release zip:" $ReleaseZip
