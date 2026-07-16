$ErrorActionPreference = 'Stop'

function Find-OutcomCodexExecutable {
    $candidates = @()

    if (-not [string]::IsNullOrWhiteSpace($env:OUTCOM_CODEX_PATH)) {
        $configuredPath = [Environment]::ExpandEnvironmentVariables($env:OUTCOM_CODEX_PATH.Trim().Trim([char]34))
        if (Test-Path -LiteralPath $configuredPath -PathType Container) {
            $configuredPath = Join-Path $configuredPath 'codex.exe'
        }

        $candidates += [pscustomobject]@{
            Path = $configuredPath
            Source = 'OUTCOM_CODEX_PATH'
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) {
        $candidates += [pscustomobject]@{
            Path = Join-Path $env:LOCALAPPDATA 'Outcom\CodexRuntime\codex.exe'
            Source = 'runtime Outcom'
        }
    }

    $pathCommand = Get-Command 'codex.exe' -CommandType Application -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if ($pathCommand) {
        $candidates += [pscustomobject]@{
            Path = $pathCommand.Source
            Source = 'PATH'
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($env:USERPROFILE)) {
        $extensionRoots = @(
            (Join-Path $env:USERPROFILE '.vscode\extensions'),
            (Join-Path $env:USERPROFILE '.vscode-insiders\extensions')
        )

        foreach ($extensionRoot in $extensionRoots) {
            if (-not (Test-Path -LiteralPath $extensionRoot -PathType Container)) {
                continue
            }

            $extensions = Get-ChildItem -LiteralPath $extensionRoot -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -like 'openai.chatgpt-*' } |
                Sort-Object -Property LastWriteTimeUtc -Descending

            foreach ($extension in $extensions) {
                foreach ($runtimeFolder in @('windows-x86_64', 'windows-arm64')) {
                    $candidates += [pscustomobject]@{
                        Path = Join-Path $extension.FullName "bin\$runtimeFolder\codex.exe"
                        Source = "extension VS Code ($($extension.Name))"
                    }
                }
            }
        }
    }

    $seenPaths = @{}
    foreach ($candidate in $candidates) {
        if ([string]::IsNullOrWhiteSpace($candidate.Path)) {
            continue
        }

        $candidateKey = $candidate.Path.ToLowerInvariant()
        if ($seenPaths.ContainsKey($candidateKey)) {
            continue
        }

        $seenPaths[$candidateKey] = $true
        if (Test-Path -LiteralPath $candidate.Path -PathType Leaf) {
            return [pscustomobject]@{
                Path = (Resolve-Path -LiteralPath $candidate.Path).Path
                Source = $candidate.Source
            }
        }
    }

    return $null
}

function Get-OutcomCodexVersion {
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
        throw "Le processus Codex n'a pas pu être démarré."
    }

    try {
        if (-not $process.WaitForExit(3000)) {
            try {
                $process.Kill()
                $null = $process.WaitForExit(1000)
            } catch {
                # L'arrêt forcé est seulement une protection du diagnostic.
            }

            return [pscustomobject]@{
                ExitCode = $null
                Text = $null
                TimedOut = $true
            }
        }

        return [pscustomobject]@{
            ExitCode = $process.ExitCode
            Text = $process.StandardOutput.ReadToEnd().Trim()
            TimedOut = $false
        }
    } finally {
        $process.Dispose()
    }
}

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

$codex = Find-OutcomCodexExecutable
$validatedCodexVersion = 'codex-cli 0.144.0-alpha.4'
if (-not $codex) {
    Write-Warning 'Codex non détecté. La compilation reste possible, mais la connexion ChatGPT d''Outcom nécessitera Codex CLI, l''extension Codex pour VS Code ou OUTCOM_CODEX_PATH.'
} else {
    $codexPath = $codex.Path
    try {
        $version = Get-OutcomCodexVersion -ExecutablePath $codexPath
        $versionExitCode = $version.ExitCode
        $versionText = $version.Text

        if ($version.TimedOut) {
            Write-Warning 'Codex a été détecté, mais la lecture de sa version a dépassé 3 secondes.'
        } elseif ($versionExitCode -eq 0 -and -not [string]::IsNullOrWhiteSpace($versionText)) {
            Write-Host "[OK] Codex : $versionText" -ForegroundColor Green
            if ($versionText -ne $validatedCodexVersion) {
                Write-Warning "La version validée pour Outcom est $validatedCodexVersion ; vérifiez la compatibilité de codex app-server avant distribution."
            }
        } else {
            Write-Warning "Codex a été détecté, mais sa version n'a pas pu être lue (code $versionExitCode)."
        }
    } catch {
        Write-Warning "Codex a été détecté, mais sa version n'a pas pu être lue : $($_.Exception.Message)"
    }

    Write-Host "     Exécutable : $codexPath ($($codex.Source))" -ForegroundColor DarkGray
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
