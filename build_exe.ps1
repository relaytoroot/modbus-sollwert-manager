param(
    [string]$PythonExe = ".\.venv\Scripts\python.exe"
)

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$distDir = Join-Path $projectRoot "dist"
$tempBuildRoot = Join-Path $env:TEMP ("FGH_Modbus_Sollwert_Manager_" + [DateTime]::Now.ToString("yyyyMMdd_HHmmss"))
$pyiDistDir = Join-Path $tempBuildRoot "pyinstaller_dist"
$pyiWorkDir = Join-Path $tempBuildRoot "pyinstaller_work"
$pyiSpecDir = Join-Path $tempBuildRoot "pyinstaller_spec"
$legacyReleaseDir = Join-Path $distDir "Release_FGH_MSC"
$releaseDir = Join-Path $distDir "Abgabe_FGH_Modbus_Sollwert_Manager"
$zipPath = Join-Path $distDir "Abgabe_FGH_Modbus_Sollwert_Manager.zip"
$exeName = "FGH_Modbus_Sollwert_Manager"
$entryScript = Join-Path $projectRoot "start_modbus_sollwert_manager.py"

Write-Host "Baue FGH Modbus Sollwert Manager..."

New-Item -ItemType Directory -Path $pyiWorkDir -Force | Out-Null
New-Item -ItemType Directory -Path $pyiSpecDir -Force | Out-Null

& $PythonExe -m PyInstaller `
    --noconfirm `
    --clean `
    --windowed `
    --name $exeName `
    --icon "$($projectRoot)\FGH_Logo_gruen.ico" `
    --distpath $pyiDistDir `
    --workpath $pyiWorkDir `
    --specpath $pyiSpecDir `
    --add-data "$($projectRoot)\FGH_Logo_prüflabor_gruen.ico;." `
    --add-data "$($projectRoot)\FGH_Logo_gruen.ico;." `
    $entryScript

if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller-Build fehlgeschlagen."
}

if (Test-Path $legacyReleaseDir) {
    Remove-Item $legacyReleaseDir -Recurse -Force
}

if (Test-Path $releaseDir) {
    Remove-Item $releaseDir -Recurse -Force
}

if (Test-Path $zipPath) {
    Remove-Item $zipPath -Force
}

New-Item -ItemType Directory -Path $releaseDir | Out-Null
Copy-Item (Join-Path $pyiDistDir "$exeName\*") $releaseDir -Recurse

$docsTarget = Join-Path $releaseDir "Dokumentation"
New-Item -ItemType Directory -Path $docsTarget | Out-Null
Copy-Item (Join-Path $projectRoot "docs\BETRIEBSANLEITUNG.md") $docsTarget

$exampleTarget = Join-Path $releaseDir "Beispieldateien"
New-Item -ItemType Directory -Path $exampleTarget | Out-Null
Copy-Item (Join-Path $projectRoot "tests\fixtures\sample_plan.xlsx") (Join-Path $exampleTarget "Beispiel_Testplan.xlsx")

Copy-Item (Join-Path $projectRoot "Start_FGH_Modbus_Sollwert_Manager.bat") $releaseDir

Compress-Archive -Path $releaseDir -DestinationPath $zipPath

if (Test-Path $tempBuildRoot) {
    Remove-Item $tempBuildRoot -Recurse -Force
}

Write-Host "Fertig. Release-Ordner: $releaseDir"
Write-Host "ZIP-Datei: $zipPath"

