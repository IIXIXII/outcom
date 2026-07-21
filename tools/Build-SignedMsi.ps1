[CmdletBinding()]
param(
    [string]$Version,
    [string]$CertificateThumbprint = $env:OUTCOM_SIGNING_CERTIFICATE_THUMBPRINT,
    [string]$ManifestCertificateThumbprint = $env:OUTCOM_MANIFEST_CERTIFICATE_THUMBPRINT,
    [string]$TimestampUrl = $(
        if ([string]::IsNullOrWhiteSpace($env:OUTCOM_TIMESTAMP_URL)) {
            'http://timestamp.digicert.com'
        } else {
            $env:OUTCOM_TIMESTAMP_URL
        }
    ),
    [switch]$SkipSigning,
    [switch]$RunIceValidation
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repositoryRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$addinProjectPath = Join-Path $repositoryRoot 'Outcom.AddIn\Outcom.AddIn.csproj'
$installerProjectPath = Join-Path $repositoryRoot 'Installer\Outcom.Installer.wixproj'
$releaseDirectory = Join-Path $repositoryRoot 'Outcom.AddIn\bin\Release'
$runtimeDirectory = Join-Path $repositoryRoot 'artifacts\CodexRuntime'
$installerArtifactDirectory = Join-Path $repositoryRoot 'artifacts\Installer'
$stageDirectory = Join-Path $installerArtifactDirectory 'stage'
$stageRuntimeDirectory = Join-Path $stageDirectory 'CodexRuntime'
$buildOutputDirectory = Join-Path $installerArtifactDirectory 'build'
$validationPath = Join-Path $installerArtifactDirectory 'outcom-msi.validation.json'
$progId = 'Outcom.AddIn'

function Assert-PathUnderRepository {
    param([Parameter(Mandatory = $true)][string]$Path)

    $resolved = [System.IO.Path]::GetFullPath($Path)
    if (-not $resolved.StartsWith(
        $repositoryRoot.TrimEnd('\') + '\',
        [StringComparison]::OrdinalIgnoreCase
    )) {
        throw "Unsafe path outside repository: $resolved"
    }

    return $resolved
}

function Find-MSBuild {
    $vswherePath = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
    if (Test-Path -LiteralPath $vswherePath) {
        $installationPath = & $vswherePath `
            -latest `
            -products '*' `
            -requires Microsoft.Component.MSBuild `
            -property installationPath
        if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace($installationPath)) {
            $candidate = Join-Path $installationPath 'MSBuild\Current\Bin\MSBuild.exe'
            if (Test-Path -LiteralPath $candidate) {
                return $candidate
            }
        }
    }

    throw 'MSBuild from Visual Studio was not found.'
}

function Find-SignTool {
    $command = Get-Command 'signtool.exe' -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if ($command) {
        return $command.Source
    }

    $roots = @(
        (Join-Path ${env:ProgramFiles(x86)} 'Windows Kits\10\bin'),
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft SDKs\ClickOnce\SignTool')
    )
    foreach ($root in $roots) {
        if (-not (Test-Path -LiteralPath $root)) {
            continue
        }

        $candidate = Get-ChildItem `
            -LiteralPath $root `
            -Filter 'signtool.exe' `
            -File `
            -Recurse `
            -ErrorAction SilentlyContinue |
            Sort-Object FullName -Descending |
            Select-Object -First 1
        if ($candidate) {
            return $candidate.FullName
        }
    }

    throw 'SignTool.exe was not found. Install the Windows SDK signing tools.'
}

function Get-CodeSigningCertificate {
    param([Parameter(Mandatory = $true)][string]$Thumbprint)

    $normalized = ($Thumbprint -replace '[^0-9A-Fa-f]', '').ToUpperInvariant()
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        throw 'The code-signing certificate thumbprint is empty.'
    }

    $certificate = Get-ChildItem -LiteralPath "Cert:\CurrentUser\My\$normalized" -ErrorAction SilentlyContinue
    if (-not $certificate) {
        throw "Certificate $normalized was not found in Cert:\CurrentUser\My."
    }
    if (-not $certificate.HasPrivateKey) {
        throw "Certificate $normalized has no accessible private key."
    }
    if ($certificate.NotBefore -gt (Get-Date) -or $certificate.NotAfter -le (Get-Date)) {
        throw "Certificate $normalized is not currently valid."
    }

    $codeSigningOid = '1.3.6.1.5.5.7.3.3'
    $supportsCodeSigning = $false
    foreach ($extension in $certificate.Extensions) {
        if ($extension.Oid.Value -ne '2.5.29.37') {
            continue
        }

        foreach ($usage in $extension.EnhancedKeyUsages) {
            if ($usage.Value -eq $codeSigningOid) {
                $supportsCodeSigning = $true
            }
        }
    }
    if (-not $supportsCodeSigning) {
        throw "Certificate $normalized is not valid for code signing."
    }

    return $certificate
}

function Find-Mage {
    $roots = @(
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft SDKs'),
        (Join-Path ${env:ProgramFiles(x86)} 'Windows Kits')
    )
    foreach ($root in $roots) {
        if (-not (Test-Path -LiteralPath $root)) {
            continue
        }

        $candidate = Get-ChildItem `
            -LiteralPath $root `
            -Filter 'mage.exe' `
            -File `
            -Recurse `
            -ErrorAction SilentlyContinue |
            Where-Object { $_.FullName -match 'NETFX 4\.8 Tools' } |
            Select-Object -First 1
        if ($candidate) {
            return $candidate.FullName
        }
    }

    throw 'Mage.exe from the .NET Framework 4.8 SDK was not found.'
}

function Get-MsiRows {
    param(
        [Parameter(Mandatory = $true)][string]$MsiPath,
        [Parameter(Mandatory = $true)][string]$Query,
        [Parameter(Mandatory = $true)][int]$FieldCount
    )

    $installer = $null
    $database = $null
    $view = $null
    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $database = $installer.GetType().InvokeMember(
            'OpenDatabase',
            'InvokeMethod',
            $null,
            $installer,
            @($MsiPath, 0)
        )
        $view = $database.GetType().InvokeMember(
            'OpenView',
            'InvokeMethod',
            $null,
            $database,
            @($Query)
        )
        $view.GetType().InvokeMember(
            'Execute',
            'InvokeMethod',
            $null,
            $view,
            $null
        ) | Out-Null

        while ($true) {
            $record = $view.GetType().InvokeMember(
                'Fetch',
                'InvokeMethod',
                $null,
                $view,
                $null
            )
            if ($null -eq $record) {
                break
            }

            $values = @()
            for ($index = 1; $index -le $FieldCount; $index++) {
                $values += [string]$record.StringData($index)
            }
            Write-Output -NoEnumerate $values
            [Runtime.InteropServices.Marshal]::FinalReleaseComObject($record) | Out-Null
        }
    }
    finally {
        if ($view) {
            [Runtime.InteropServices.Marshal]::FinalReleaseComObject($view) | Out-Null
        }
        if ($database) {
            [Runtime.InteropServices.Marshal]::FinalReleaseComObject($database) | Out-Null
        }
        if ($installer) {
            [Runtime.InteropServices.Marshal]::FinalReleaseComObject($installer) | Out-Null
        }
    }
}

function Test-MsiDatabase {
    param(
        [Parameter(Mandatory = $true)][string]$MsiPath,
        [Parameter(Mandatory = $true)][string]$ExpectedVersion
    )

    $properties = @{}
    $propertyRows = @(Get-MsiRows `
        -MsiPath $MsiPath `
        -Query 'SELECT `Property`, `Value` FROM `Property`' `
        -FieldCount 2)
    foreach ($row in $propertyRows) {
        $properties[$row[0]] = $row[1]
    }
    if ($properties['ProductName'] -ne 'Outcom' -or
        $properties['ProductVersion'] -ne $ExpectedVersion -or
        $properties['ALLUSERS'] -ne '1') {
        throw 'The MSI product identity, version, or per-machine scope is invalid.'
    }

    $fileNames = @(Get-MsiRows `
        -MsiPath $MsiPath `
        -Query 'SELECT `FileName` FROM `File`' `
        -FieldCount 1 | ForEach-Object { $_[0] })
    foreach ($requiredFileName in @('Outcom.AddIn.dll', 'Outcom.AddIn.vsto', 'codex.exe')) {
        if (-not ($fileNames | Where-Object {
            $_ -eq $requiredFileName -or $_ -like "*|$requiredFileName"
        })) {
            throw "The MSI File table is missing $requiredFileName."
        }
    }

    $registryRows = @(Get-MsiRows `
        -MsiPath $MsiPath `
        -Query 'SELECT `Key`, `Name`, `Value` FROM `Registry`' `
        -FieldCount 3)
    $manifestRow = $registryRows | Where-Object {
        $_[0] -eq 'Software\Microsoft\Office\Outlook\Addins\Outcom.AddIn' -and
        $_[1] -eq 'Manifest'
    } | Select-Object -First 1
    $loadBehaviorRow = $registryRows | Where-Object {
        $_[0] -eq 'Software\Microsoft\Office\Outlook\Addins\Outcom.AddIn' -and
        $_[1] -eq 'LoadBehavior'
    } | Select-Object -First 1
    if (-not $manifestRow -or
        $manifestRow[2] -ne 'file:///[INSTALLFOLDER]Outcom.AddIn.vsto|vstolocal' -or
        -not $loadBehaviorRow -or
        $loadBehaviorRow[2] -ne '#3') {
        throw 'The MSI Outlook registration is invalid.'
    }
}

function Copy-DirectoryFiles {
    param(
        [Parameter(Mandatory = $true)][string]$Source,
        [Parameter(Mandatory = $true)][string]$Destination,
        [string[]]$ExcludedExtensions = @()
    )

    $sourceRoot = [System.IO.Path]::GetFullPath($Source).TrimEnd('\')
    foreach ($file in Get-ChildItem -LiteralPath $sourceRoot -File -Recurse) {
        if ($ExcludedExtensions -contains $file.Extension.ToLowerInvariant()) {
            continue
        }

        $relativePath = $file.FullName.Substring($sourceRoot.Length).TrimStart('\')
        $destinationPath = Join-Path $Destination $relativePath
        $destinationParent = Split-Path -Parent $destinationPath
        New-Item -ItemType Directory -Path $destinationParent -Force | Out-Null
        Copy-Item -LiteralPath $file.FullName -Destination $destinationPath -Force
    }
}

if (-not (Test-Path -LiteralPath $addinProjectPath) -or
    -not (Test-Path -LiteralPath $installerProjectPath)) {
    throw 'The Outcom project or installer project is missing.'
}

if ([string]::IsNullOrWhiteSpace($Version)) {
    [xml]$addinProject = Get-Content -LiteralPath $addinProjectPath -Raw -Encoding UTF8
    $versionNode = Select-Xml `
        -Xml $addinProject `
        -XPath "/*[local-name()='Project']/*[local-name()='PropertyGroup']/*[local-name()='OutcomVersion']" |
        Select-Object -First 1
    if (-not $versionNode) {
        throw 'OutcomVersion was not found in the VSTO project.'
    }
    $Version = [string]$versionNode.Node.InnerText
}
if ($Version -notmatch '^(?<Major>[0-9]+)\.(?<Minor>[0-9]+)\.(?<Build>[0-9]+)$') {
    throw "MSI version must contain three numeric fields: $Version"
}
if ([int]$Matches.Major -gt 255 -or
    [int]$Matches.Minor -gt 255 -or
    [int]$Matches.Build -gt 65535) {
    throw "MSI version is outside Windows Installer limits: $Version"
}

$signingCertificate = $null
if (-not $SkipSigning) {
    if ([string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
        throw 'Set OUTCOM_SIGNING_CERTIFICATE_THUMBPRINT or pass -CertificateThumbprint. Use -SkipSigning only for local structural validation.'
    }
    if ([string]::IsNullOrWhiteSpace($TimestampUrl)) {
        throw 'An RFC 3161 timestamp URL is required for a distribution signature.'
    }

    $signingCertificate = Get-CodeSigningCertificate -Thumbprint $CertificateThumbprint
    $CertificateThumbprint = $signingCertificate.Thumbprint
    if ([string]::IsNullOrWhiteSpace($ManifestCertificateThumbprint)) {
        $ManifestCertificateThumbprint = $CertificateThumbprint
    }
    $null = Get-CodeSigningCertificate -Thumbprint $ManifestCertificateThumbprint
}

$msbuildPath = Find-MSBuild
Write-Host "Building Outcom $Version Release..."
$addinBuildArguments = @(
    $addinProjectPath,
    '/t:Rebuild',
    '/p:Configuration=Release',
    '/p:Platform=AnyCPU',
    "/p:OutcomVersion=$Version",
    '/p:DefineConstants=VSTO40%3BUSEOFFICEINTEROP%3BTRACE',
    '/v:minimal'
)
if (-not [string]::IsNullOrWhiteSpace($ManifestCertificateThumbprint)) {
    $addinBuildArguments += "/p:OUTCOM_MANIFEST_CERTIFICATE_THUMBPRINT=$ManifestCertificateThumbprint"
    $addinBuildArguments += '/p:OUTCOM_MANIFEST_KEY_FILE='
}
& $msbuildPath @addinBuildArguments
if ($LASTEXITCODE -ne 0) {
    throw "Outcom Release build failed with exit code $LASTEXITCODE."
}

$requiredReleaseFiles = @(
    'Outcom.AddIn.dll',
    'Outcom.AddIn.dll.config',
    'Outcom.AddIn.dll.manifest',
    'Outcom.AddIn.vsto'
)
foreach ($requiredFile in $requiredReleaseFiles) {
    if (-not (Test-Path -LiteralPath (Join-Path $releaseDirectory $requiredFile) -PathType Leaf)) {
        throw "Release output is missing $requiredFile."
    }
}

$magePath = Find-Mage
& $magePath -Verify (Join-Path $releaseDirectory 'Outcom.AddIn.vsto')
if ($LASTEXITCODE -ne 0) {
    throw 'The VSTO deployment manifest signature is invalid.'
}
& $magePath -Verify (Join-Path $releaseDirectory 'Outcom.AddIn.dll.manifest')
if ($LASTEXITCODE -ne 0) {
    throw 'The VSTO application manifest signature is invalid.'
}

Write-Host 'Validating the pinned Codex runtime...'
& (Join-Path $PSScriptRoot 'Prepare-CodexRuntime.ps1') -OutputDirectory $runtimeDirectory
if ($LASTEXITCODE -ne 0) {
    throw "Codex runtime validation failed with exit code $LASTEXITCODE."
}

$requiredRuntimeFiles = @(
    'codex.exe',
    'CodexRuntime.json',
    'codex-runtime.validation.json',
    'Codex-CLI-LICENSE.txt'
)
foreach ($requiredFile in $requiredRuntimeFiles) {
    if (-not (Test-Path -LiteralPath (Join-Path $runtimeDirectory $requiredFile) -PathType Leaf)) {
        throw "Codex runtime output is missing $requiredFile."
    }
}

$stageDirectory = Assert-PathUnderRepository -Path $stageDirectory
$buildOutputDirectory = Assert-PathUnderRepository -Path $buildOutputDirectory
if (Test-Path -LiteralPath $stageDirectory) {
    Remove-Item -LiteralPath $stageDirectory -Recurse -Force
}
if (Test-Path -LiteralPath $buildOutputDirectory) {
    Remove-Item -LiteralPath $buildOutputDirectory -Recurse -Force
}
New-Item -ItemType Directory -Path $stageDirectory -Force | Out-Null
New-Item -ItemType Directory -Path $stageRuntimeDirectory -Force | Out-Null
New-Item -ItemType Directory -Path $buildOutputDirectory -Force | Out-Null

Copy-DirectoryFiles `
    -Source $releaseDirectory `
    -Destination $stageDirectory `
    -ExcludedExtensions @('.pdb')
Copy-DirectoryFiles -Source $runtimeDirectory -Destination $stageRuntimeDirectory
Copy-Item -LiteralPath (Join-Path $repositoryRoot 'LICENSE') -Destination $stageDirectory -Force
Copy-Item -LiteralPath (Join-Path $repositoryRoot 'THIRD-PARTY-NOTICES.md') -Destination $stageDirectory -Force

Write-Host 'Building the x64 MSI with WiX Toolset 5.0.2...'
& dotnet build $installerProjectPath `
    --configuration Release `
    --output $buildOutputDirectory `
    "/p:OutcomVersion=$Version" `
    "/p:StageDirectory=$stageDirectory" `
    "/p:SuppressValidation=$(-not $RunIceValidation)" `
    --nologo
if ($LASTEXITCODE -ne 0) {
    throw "WiX build failed with exit code $LASTEXITCODE."
}

$builtMsiPath = Join-Path $buildOutputDirectory "Outcom-$Version-x64.msi"
if (-not (Test-Path -LiteralPath $builtMsiPath -PathType Leaf)) {
    $builtMsiPath = Get-ChildItem -LiteralPath $buildOutputDirectory -Filter '*.msi' -File -Recurse |
        Select-Object -First 1 -ExpandProperty FullName
}
if ([string]::IsNullOrWhiteSpace($builtMsiPath) -or
    -not (Test-Path -LiteralPath $builtMsiPath -PathType Leaf)) {
    throw 'WiX did not produce an MSI package.'
}

$finalMsiName = if ($SkipSigning) {
    "Outcom-$Version-x64-unsigned.msi"
} else {
    "Outcom-$Version-x64.msi"
}
$finalMsiPath = Join-Path $installerArtifactDirectory $finalMsiName
Copy-Item -LiteralPath $builtMsiPath -Destination $finalMsiPath -Force

Test-MsiDatabase -MsiPath $finalMsiPath -ExpectedVersion $Version

$signatureStatus = 'NotSigned'
if (-not $SkipSigning) {
    $signToolPath = Find-SignTool
    Write-Host "Signing the MSI with $($signingCertificate.Subject)..."
    & $signToolPath sign `
        /sha1 $CertificateThumbprint `
        /s My `
        /fd SHA256 `
        /tr $TimestampUrl `
        /td SHA256 `
        /d 'Outcom Outlook Add-in' `
        $finalMsiPath
    if ($LASTEXITCODE -ne 0) {
        throw "MSI signing failed with exit code $LASTEXITCODE."
    }

    & $signToolPath verify /pa /all /v $finalMsiPath
    if ($LASTEXITCODE -ne 0) {
        throw "MSI signature verification failed with exit code $LASTEXITCODE."
    }
    $signatureStatus = 'Valid'
}

$msiFile = Get-Item -LiteralPath $finalMsiPath
$msiHash = (Get-FileHash -LiteralPath $finalMsiPath -Algorithm SHA256).Hash.ToLowerInvariant()
$hashPath = $finalMsiPath + '.sha256'
Set-Content `
    -LiteralPath $hashPath `
    -Value ($msiHash + '  ' + $msiFile.Name) `
    -Encoding ASCII

$validation = [ordered]@{
    schemaVersion = 1
    validatedAtUtc = [DateTime]::UtcNow.ToString('o')
    product = 'Outcom'
    version = $Version
    platform = 'x64'
    scope = 'perMachine'
    progId = $progId
    msi = $msiFile.Name
    size = $msiFile.Length
    sha256 = $msiHash
    signatureStatus = $signatureStatus
    signingCertificateThumbprint = if ($signingCertificate) {
        $signingCertificate.Thumbprint
    } else {
        $null
    }
    signingCertificateSubject = if ($signingCertificate) {
        $signingCertificate.Subject
    } else {
        $null
    }
    timestampUrl = if ($SkipSigning) { $null } else { $TimestampUrl }
    manifestCertificateThumbprint = if ([string]::IsNullOrWhiteSpace($ManifestCertificateThumbprint)) {
        'development certificate from project'
    } else {
        $ManifestCertificateThumbprint
    }
    codexRuntimeValidation = 'passed'
    msiDatabaseValidation = 'passed'
    iceValidation = if ($RunIceValidation) { 'passed' } else { 'not-run' }
}
$validation |
    ConvertTo-Json -Depth 4 |
    Set-Content -LiteralPath $validationPath -Encoding UTF8

if ($SkipSigning) {
    Write-Warning 'The MSI is unsigned and uses the development manifest certificate. Do not distribute it.'
}
Write-Host "[OK] MSI: $finalMsiPath" -ForegroundColor Green
Write-Host "[OK] SHA-256: $msiHash" -ForegroundColor Green
Write-Host "[OK] Validation: $validationPath" -ForegroundColor Green
