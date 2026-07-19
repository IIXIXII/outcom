[CmdletBinding()]
param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\artifacts\CodexRuntime'),
    [switch] $InstallForCurrentUser
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-Sha256 {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

function Assert-Sha256 {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path,

        [Parameter(Mandatory = $true)]
        [string] $Expected,

        [Parameter(Mandatory = $true)]
        [string] $Description
    )

    $actual = Get-Sha256 -Path $Path
    if ($actual -ne $Expected.ToLowerInvariant()) {
        throw "$Description ne correspond pas à l'empreinte SHA-256 épinglée. Attendue : $Expected ; obtenue : $actual."
    }
}

function Get-PeMachine {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    $stream = [System.IO.File]::Open(
        $Path,
        [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::Read,
        [System.IO.FileShare]::Read
    )
    try {
        $reader = New-Object System.IO.BinaryReader($stream)
        try {
            if ($reader.ReadUInt16() -ne 0x5A4D) {
                throw "Le runtime Codex n'est pas un exécutable Windows PE valide."
            }

            $stream.Position = 0x3C
            $peOffset = $reader.ReadInt32()
            if ($peOffset -lt 0 -or $peOffset -gt ($stream.Length - 6)) {
                throw "L'en-tête PE du runtime Codex est invalide."
            }

            $stream.Position = $peOffset
            if ($reader.ReadUInt32() -ne 0x00004550) {
                throw "La signature PE du runtime Codex est invalide."
            }

            return ('0x{0:X4}' -f $reader.ReadUInt16())
        } finally {
            $reader.Dispose()
        }
    } finally {
        $stream.Dispose()
    }
}

function ConvertTo-WindowsCommandLineArgument {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string] $Value
    )

    if ($Value.Length -gt 0 -and $Value.IndexOfAny([char[]]@(' ', "`t", "`n", "`v", '"')) -lt 0) {
        return $Value
    }

    $result = New-Object System.Text.StringBuilder
    $null = $result.Append('"')
    $backslashCount = 0
    foreach ($character in $Value.ToCharArray()) {
        if ($character -eq '\') {
            $backslashCount++
            continue
        }

        if ($character -eq '"') {
            $null = $result.Append('\', ($backslashCount * 2) + 1)
            $null = $result.Append('"')
            $backslashCount = 0
            continue
        }

        $null = $result.Append('\', $backslashCount)
        $backslashCount = 0
        $null = $result.Append($character)
    }

    $null = $result.Append('\', $backslashCount * 2)
    $null = $result.Append('"')
    return $result.ToString()
}

function Invoke-CodexVersion {
    param(
        [Parameter(Mandatory = $true)]
        [string] $ExecutablePath
    )

    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = $ExecutablePath
    $startInfo.Arguments = '--version'
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true

    $process = [System.Diagnostics.Process]::Start($startInfo)
    if (-not $process) {
        throw "Le runtime Codex n'a pas pu être démarré."
    }

    try {
        $outputTask = $process.StandardOutput.ReadToEndAsync()
        $errorTask = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit(10000)) {
            try {
                $process.Kill()
            } catch {
            }

            throw "La lecture de la version Codex a dépassé dix secondes."
        }

        if (-not [System.Threading.Tasks.Task]::WaitAll(
            [System.Threading.Tasks.Task[]]@($outputTask, $errorTask),
            2000
        )) {
            throw "La sortie de version Codex n'a pas pu être lue."
        }

        if ($process.ExitCode -ne 0) {
            throw "La commande codex --version a échoué avec le code $($process.ExitCode)."
        }

        return $outputTask.Result.Trim()
    } finally {
        $process.Dispose()
    }
}

function Test-CodexAppServerHandshake {
    param(
        [Parameter(Mandatory = $true)]
        [string] $ExecutablePath,

        [Parameter(Mandatory = $true)]
        [string] $TemporaryRoot
    )

    $codexHome = Join-Path $TemporaryRoot 'codex-home'
    $workspace = Join-Path $TemporaryRoot 'workspace'
    $null = New-Item -ItemType Directory -Path $codexHome -Force
    $null = New-Item -ItemType Directory -Path $workspace -Force

    $arguments = @(
        'app-server',
        '--stdio',
        '--strict-config',
        '-c',
        'web_search="disabled"',
        '-c',
        'cli_auth_credentials_store="keyring"',
        '-c',
        'forced_login_method="chatgpt"',
        '-c',
        'mcp_servers={}'
    )
    foreach ($feature in @(
        'shell_tool',
        'unified_exec',
        'apps',
        'browser_use',
        'computer_use',
        'image_generation',
        'in_app_browser',
        'multi_agent',
        'hooks',
        'plugins',
        'remote_plugin',
        'skill_mcp_dependency_install',
        'tool_call_mcp_elicitation'
    )) {
        $arguments += '--disable'
        $arguments += $feature
    }

    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = $ExecutablePath
    $startInfo.Arguments = (($arguments | ForEach-Object {
        ConvertTo-WindowsCommandLineArgument -Value $_
    }) -join ' ')
    $startInfo.WorkingDirectory = $workspace
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardInput = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.StandardOutputEncoding = New-Object System.Text.UTF8Encoding($false)
    $startInfo.StandardErrorEncoding = New-Object System.Text.UTF8Encoding($false)
    $startInfo.EnvironmentVariables['CODEX_HOME'] = $codexHome

    foreach ($variableName in @(
        'OPENAI_API_KEY',
        'OPENAI_BASE_URL',
        'OPENAI_ORGANIZATION',
        'OPENAI_PROJECT',
        'AZURE_OPENAI_API_KEY',
        'AZURE_OPENAI_ENDPOINT',
        'CODEX_ACCESS_TOKEN',
        'CODEX_API_KEY',
        'AWS_ACCESS_KEY_ID',
        'AWS_SECRET_ACCESS_KEY',
        'AWS_SESSION_TOKEN'
    )) {
        $startInfo.EnvironmentVariables.Remove($variableName)
    }

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
    if (-not $process.Start()) {
        $process.Dispose()
        throw "codex app-server n'a pas pu être démarré."
    }

    $errorDrainTask = $process.StandardError.ReadToEndAsync()
    try {
        $initialize = @{
            id = 1
            method = 'initialize'
            params = @{
                clientInfo = @{
                    name = 'outcom-runtime-validation'
                    title = 'Outcom runtime validation'
                    version = '0.1.0'
                }
                capabilities = @{
                    experimentalApi = $false
                    requestAttestation = $false
                    mcpServerOpenaiFormElicitation = $false
                }
            }
        } | ConvertTo-Json -Compress -Depth 8

        $utf8 = New-Object System.Text.UTF8Encoding($false)
        $bytes = $utf8.GetBytes($initialize + "`n")
        $process.StandardInput.BaseStream.Write($bytes, 0, $bytes.Length)
        $process.StandardInput.BaseStream.Flush()

        $deadline = [DateTime]::UtcNow.AddSeconds(20)
        $response = $null
        while ([DateTime]::UtcNow -lt $deadline -and -not $response) {
            $remaining = [int][Math]::Max(
                1,
                ($deadline - [DateTime]::UtcNow).TotalMilliseconds
            )
            $readTask = $process.StandardOutput.ReadLineAsync()
            if (-not $readTask.Wait($remaining)) {
                break
            }

            $line = $readTask.Result
            if ([string]::IsNullOrWhiteSpace($line)) {
                if ($process.HasExited) {
                    break
                }

                continue
            }

            $message = $line | ConvertFrom-Json
            if ($message.PSObject.Properties.Name -contains 'id' -and
                [string]$message.id -eq '1') {
                $response = $message
            }
        }

        if (-not $response) {
            throw "codex app-server n'a pas répondu à initialize dans le délai imparti."
        }

        if ($response.PSObject.Properties.Name -contains 'error') {
            throw "codex app-server a refusé initialize."
        }

        $returnedHome = [System.IO.Path]::GetFullPath([string]$response.result.codexHome)
        $expectedHome = [System.IO.Path]::GetFullPath($codexHome)
        if ($returnedHome.TrimEnd('\') -ne $expectedHome.TrimEnd('\')) {
            throw "codex app-server n'a pas confirmé le profil isolé demandé."
        }

        $initialized = '{"method":"initialized","params":null}' + "`n"
        $bytes = $utf8.GetBytes($initialized)
        $process.StandardInput.BaseStream.Write($bytes, 0, $bytes.Length)
        $process.StandardInput.BaseStream.Flush()
    } finally {
        try {
            $process.StandardInput.Close()
        } catch {
        }

        if (-not $process.HasExited) {
            try {
                $process.Kill()
                $null = $process.WaitForExit(3000)
            } catch {
            }
        }

        try {
            $null = $errorDrainTask.Wait(2000)
        } catch {
        }

        $process.Dispose()
    }
}

$manifestPath = [System.IO.Path]::GetFullPath(
    (Join-Path $PSScriptRoot '..\Outcom.AddIn\CodexRuntime.json')
)
$manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ($manifest.schemaVersion -ne 1) {
    throw "La version du manifeste CodexRuntime.json n'est pas prise en charge."
}

if (-not [Environment]::Is64BitOperatingSystem) {
    throw "Le runtime Codex épinglé par Outcom nécessite Windows x64."
}

if ($InstallForCurrentUser) {
    if ([string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) {
        throw 'LOCALAPPDATA est indisponible.'
    }

    $destination = Join-Path $env:LOCALAPPDATA 'Outcom\CodexRuntime'
} else {
    $destination = $OutputDirectory
}
$destination = [System.IO.Path]::GetFullPath($destination)
$null = New-Item -ItemType Directory -Path $destination -Force

$temporaryBase = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath()).TrimEnd('\')
$temporaryRoot = Join-Path $temporaryBase ('Outcom-CodexRuntime-' + [Guid]::NewGuid().ToString('N'))
$temporaryRoot = [System.IO.Path]::GetFullPath($temporaryRoot)
if (-not $temporaryRoot.StartsWith(
    $temporaryBase + '\',
    [StringComparison]::OrdinalIgnoreCase
)) {
    throw "Le dossier temporaire du runtime Codex est invalide."
}

$null = New-Item -ItemType Directory -Path $temporaryRoot
try {
    $targetExecutable = Join-Path $destination 'codex.exe'
    $useExistingExecutable = $false
    if (Test-Path -LiteralPath $targetExecutable -PathType Leaf) {
        try {
            Assert-Sha256 `
                -Path $targetExecutable `
                -Expected $manifest.executableSha256 `
                -Description 'Le runtime Codex existant'
            $useExistingExecutable = $true
            Write-Host '[OK] Binaire Codex déjà présent avec la bonne empreinte.' -ForegroundColor Green
        } catch {
            Write-Warning 'Le runtime Codex existant sera remplacé par la version épinglée.'
        }
    }

    if (-not $useExistingExecutable) {
        $archivePath = Join-Path $temporaryRoot $manifest.assetName
        Write-Host "Téléchargement de Codex $($manifest.version) depuis la release officielle OpenAI…"
        Invoke-WebRequest `
            -Uri $manifest.downloadUrl `
            -OutFile $archivePath `
            -UseBasicParsing `
            -Headers @{ 'User-Agent' = 'Outcom-Runtime-Preparation' }
        Assert-Sha256 `
            -Path $archivePath `
            -Expected $manifest.assetSha256 `
            -Description "L'archive Codex téléchargée"
        Write-Host "[OK] Empreinte SHA-256 de l'archive validée." -ForegroundColor Green

        $extractPath = Join-Path $temporaryRoot 'extracted'
        Expand-Archive -LiteralPath $archivePath -DestinationPath $extractPath
        $extractedExecutable = Join-Path $extractPath $manifest.executableName
        if (-not (Test-Path -LiteralPath $extractedExecutable -PathType Leaf)) {
            throw "L'archive officielle ne contient pas $($manifest.executableName)."
        }

        Assert-Sha256 `
            -Path $extractedExecutable `
            -Expected $manifest.executableSha256 `
            -Description 'Le binaire Codex extrait'
        Copy-Item -LiteralPath $extractedExecutable -Destination $targetExecutable -Force
    }

    Assert-Sha256 `
        -Path $targetExecutable `
        -Expected $manifest.executableSha256 `
        -Description 'Le binaire Codex préparé'

    $peMachine = Get-PeMachine -Path $targetExecutable
    if ($peMachine -ne $manifest.peMachine) {
        throw "Architecture PE inattendue : $peMachine au lieu de $($manifest.peMachine)."
    }
    Write-Host "[OK] Architecture Windows x64 validée ($peMachine)." -ForegroundColor Green

    $reportedVersion = Invoke-CodexVersion -ExecutablePath $targetExecutable
    if ($reportedVersion -ne $manifest.cliVersion) {
        throw "Version Codex inattendue : $reportedVersion au lieu de $($manifest.cliVersion)."
    }
    Write-Host "[OK] Version validée : $reportedVersion." -ForegroundColor Green

    Test-CodexAppServerHandshake `
        -ExecutablePath $targetExecutable `
        -TemporaryRoot $temporaryRoot
    Write-Host '[OK] Handshake app-server sécurisé validé.' -ForegroundColor Green

    $signature = Get-AuthenticodeSignature -LiteralPath $targetExecutable
    if ($signature.Status -ne [System.Management.Automation.SignatureStatus]::Valid -or
        -not $signature.SignerCertificate -or
        $signature.SignerCertificate.Subject -notlike
            "*$($manifest.authenticodePublisher)*") {
        throw "La signature Authenticode du runtime Codex n'est pas valide ou ne provient pas de $($manifest.authenticodePublisher)."
    }
    Write-Host "[OK] Signature Authenticode validée : $($manifest.authenticodePublisher)." -ForegroundColor Green

    $licensePath = Join-Path $destination 'Codex-CLI-LICENSE.txt'
    $licenseIsValid = $false
    if (Test-Path -LiteralPath $licensePath -PathType Leaf) {
        $licenseIsValid = (Get-Sha256 -Path $licensePath) -eq $manifest.licenseSha256
    }
    if (-not $licenseIsValid) {
        $temporaryLicense = Join-Path $temporaryRoot 'Codex-CLI-LICENSE.txt'
        Invoke-WebRequest `
            -Uri $manifest.licenseUrl `
            -OutFile $temporaryLicense `
            -UseBasicParsing `
            -Headers @{ 'User-Agent' = 'Outcom-Runtime-Preparation' }
        Assert-Sha256 `
            -Path $temporaryLicense `
            -Expected $manifest.licenseSha256 `
            -Description 'La licence Codex CLI'
        Copy-Item -LiteralPath $temporaryLicense -Destination $licensePath -Force
    }
    Write-Host '[OK] Licence Apache 2.0 incluse.' -ForegroundColor Green

    Copy-Item `
        -LiteralPath $manifestPath `
        -Destination (Join-Path $destination 'CodexRuntime.json') `
        -Force

    $validation = [ordered]@{
        schemaVersion = 1
        validatedAtUtc = [DateTime]::UtcNow.ToString('o')
        cliVersion = $reportedVersion
        platform = $manifest.platform
        executableSha256 = Get-Sha256 -Path $targetExecutable
        peMachine = $peMachine
        appServerHandshake = 'passed'
        authenticodeStatus = [string]$signature.Status
        authenticodeSubject = $signature.SignerCertificate.Subject
        authenticodeThumbprint = $signature.SignerCertificate.Thumbprint
        source = $manifest.downloadUrl
    }
    $validation |
        ConvertTo-Json -Depth 4 |
        Set-Content `
            -LiteralPath (Join-Path $destination 'codex-runtime.validation.json') `
            -Encoding UTF8

    Write-Host "[OK] Runtime Codex prêt : $destination" -ForegroundColor Green
} finally {
    $resolvedTemporaryRoot = [System.IO.Path]::GetFullPath($temporaryRoot)
    if ($resolvedTemporaryRoot.StartsWith(
        $temporaryBase + '\Outcom-CodexRuntime-',
        [StringComparison]::OrdinalIgnoreCase
    ) -and (Test-Path -LiteralPath $resolvedTemporaryRoot)) {
        Remove-Item -LiteralPath $resolvedTemporaryRoot -Recurse -Force
    }
}
