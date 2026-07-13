$ErrorActionPreference = 'Stop'

Write-Host 'Outcom - vérification du poste de développement' -ForegroundColor Cyan

$officeConfigurationPath = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
if (Test-Path $officeConfigurationPath) {
    $office = Get-ItemProperty $officeConfigurationPath
    Write-Host "[OK] Office : $($office.ProductReleaseIds) $($office.VersionToReport) $($office.Platform)" -ForegroundColor Green
} else {
    Write-Warning 'Office Click-to-Run non détecté.'
}

$outlookPath = 'C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE'
if (Test-Path $outlookPath) {
    Write-Host "[OK] Outlook : $outlookPath" -ForegroundColor Green
} else {
    Write-Warning 'Outlook 64 bits non détecté à son emplacement habituel.'
}

$vswherePath = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
if (-not (Test-Path $vswherePath)) {
    Write-Warning 'Visual Studio Installer non détecté. Installez Visual Studio 2022 avec la charge Développement Office/SharePoint.'
    exit 1
}

$visualStudio = & $vswherePath -latest -products * -requires Microsoft.VisualStudio.Component.Sharepoint.Tools -format json | ConvertFrom-Json
if (-not $visualStudio) {
    Write-Warning 'Visual Studio est présent, mais les outils de développement Office ne sont pas installés.'
    exit 1
}

Write-Host "[OK] Visual Studio : $($visualStudio.displayName) $($visualStudio.catalog.productDisplayVersion)" -ForegroundColor Green
Write-Host '[OK] Outils de développement Office installés.' -ForegroundColor Green
