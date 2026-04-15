param(
    [string]$PythonExe = ".\.venv\Scripts\python.exe"
)

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$distDir = Join-Path $projectRoot "dist"
$buildDir = Join-Path $projectRoot "build"
$pyiDistDir = Join-Path $distDir "pyinstaller_dist"
$pyiWorkDir = Join-Path $buildDir "pyinstaller_work"
$pyiSpecDir = Join-Path $buildDir "pyinstaller_spec"
$legacyReleaseDir = Join-Path $distDir "Release_FGH_MSC"
$releaseDir = Join-Path $distDir "Release_FGH_Modbus_Sollwert_Manager"
$exeName = "FGH_Modbus_Sollwert_Manager"
$entryScript = Join-Path $projectRoot "start_modbus_sollwert_manager.py"

Write-Host "Baue FGH Modbus Sollwert Manager..."

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
    --add-data "$($projectRoot)\combo_arrow_dark.svg;." `
    --add-data "$($projectRoot)\combo_arrow_light.svg;." `
    --add-data "$($projectRoot)\spin_arrow_up_dark.svg;." `
    --add-data "$($projectRoot)\spin_arrow_up_light.svg;." `
    --add-data "$($projectRoot)\spin_arrow_down_dark.svg;." `
    --add-data "$($projectRoot)\spin_arrow_down_light.svg;." `
    --add-data "$($projectRoot)\checkbox_check_dark.svg;." `
    --add-data "$($projectRoot)\checkbox_check_light.svg;." `
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

New-Item -ItemType Directory -Path $releaseDir | Out-Null
Copy-Item (Join-Path $pyiDistDir "$exeName\*") $releaseDir -Recurse
Copy-Item (Join-Path $projectRoot "README.md") $releaseDir
Copy-Item (Join-Path $projectRoot "UEBERGABE_MODBUS.md") $releaseDir

$docsTarget = Join-Path $releaseDir "Dokumentation"
New-Item -ItemType Directory -Path $docsTarget | Out-Null
Copy-Item (Join-Path $projectRoot "docs\BETRIEBSANLEITUNG.md") $docsTarget

$exampleTarget = Join-Path $releaseDir "Beispieldateien"
New-Item -ItemType Directory -Path $exampleTarget | Out-Null
Copy-Item (Join-Path $projectRoot "tests\fixtures\sample_plan.xlsx") (Join-Path $exampleTarget "Beispiel_Testplan.xlsx")

Copy-Item (Join-Path $projectRoot "Start_FGH_Modbus_Sollwert_Manager.bat") $releaseDir

Write-Host "Fertig. Release-Ordner: $releaseDir"

