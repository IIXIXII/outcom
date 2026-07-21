[CmdletBinding()]
param(
    [ValidateSet('Prepare', 'Status', 'Restore')]
    [string]$Action = 'Status',

    [ValidateRange(1100, 15000)]
    [int]$StartupDelayMilliseconds = 2000
)

$ErrorActionPreference = 'Stop'

$progId = 'Outcom.AddIn'
$addinRegistryPath = 'HKCU:\Software\Microsoft\Office\Outlook\Addins\Outcom.AddIn'
$resiliencyRegistryPath = 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency'
$doNotDisableRegistryPath = Join-Path $resiliencyRegistryPath 'DoNotDisableAddinList'
$managedAddinsRegistryPath = Join-Path $resiliencyRegistryPath 'AddinList'
$policyAddinsRegistryPath = 'HKCU:\Software\Policies\Microsoft\Office\16.0\Outlook\Resiliency\AddinList'
$diagnosticsRegistryPath = 'HKCU:\Software\Outcom\Diagnostics'
$startupDelayValueName = 'StartupDelayMilliseconds'
$stateDirectoryPath = Join-Path $env:LOCALAPPDATA 'Outcom\Diagnostics'
$stateFilePath = Join-Path $stateDirectoryPath 'outlook-resiliency-test-state.json'

function Get-RegistryValueSnapshot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return [pscustomobject]@{
            Exists = $false
            Kind = $null
            Value = $null
        }
    }

    $key = Get-Item -LiteralPath $Path
    if ($key.GetValueNames() -notcontains $Name) {
        return [pscustomobject]@{
            Exists = $false
            Kind = $null
            Value = $null
        }
    }

    return [pscustomobject]@{
        Exists = $true
        Kind = $key.GetValueKind($Name).ToString()
        Value = $key.GetValue($Name, $null, 'DoNotExpandEnvironmentNames')
    }
}

function Set-RegistryValueFromSnapshot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [object]$Snapshot
    )

    if ($Snapshot.Exists) {
        New-Item -Path $Path -Force | Out-Null
        New-ItemProperty `
            -LiteralPath $Path `
            -Name $Name `
            -Value $Snapshot.Value `
            -PropertyType $Snapshot.Kind `
            -Force | Out-Null
        return
    }

    if (Test-Path -LiteralPath $Path) {
        Remove-ItemProperty -LiteralPath $Path -Name $Name -ErrorAction SilentlyContinue
    }
}

function Assert-OutlookStopped {
    if (Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue) {
        throw 'Close Outlook completely before preparing or restoring the test.'
    }
}

function Get-OutcomStartupEvents {
    param(
        [Nullable[datetime]]$Since = $null,
        [int]$MaximumEvents = 100
    )

    $filter = @{
        LogName = 'Application'
        Id = 45
    }
    if ($Since.HasValue) {
        $filter.StartTime = $Since.Value
    }

    $events = Get-WinEvent -FilterHashtable $filter -MaxEvents $MaximumEvents -ErrorAction SilentlyContinue |
        Where-Object {
            $_.ProviderName -eq 'Outlook' -and
            $_.Message -match [regex]::Escape($progId)
        }

    foreach ($event in $events) {
        $addinBlockMatch = [regex]::Match(
            $event.Message,
            '(?is)ProgID[^\r\n]*' + [regex]::Escape($progId) + '(?<Block>.*?)(?=(?:\r?\n){2}|\z)')
        if (-not $addinBlockMatch.Success) {
            continue
        }

        $durationMatch = [regex]::Match(
            $addinBlockMatch.Groups['Block'].Value,
            '(?im)^[^\r\n]*millisecond[^0-9]*(?<Duration>[0-9]+)\s*$')
        if (-not $durationMatch.Success) {
            continue
        }

        [pscustomobject]@{
            TimeCreated = $event.TimeCreated
            DurationMilliseconds = [int]$durationMatch.Groups['Duration'].Value
            RecordId = $event.RecordId
        }
    }
}

function Get-OutcomDisableEvents {
    param(
        [Nullable[datetime]]$Since = $null,
        [int]$MaximumEvents = 100
    )

    $filter = @{
        LogName = 'Application'
        Id = 59
    }
    if ($Since.HasValue) {
        $filter.StartTime = $Since.Value
    }

    @(Get-WinEvent -FilterHashtable $filter -MaxEvents $MaximumEvents -ErrorAction SilentlyContinue |
        Where-Object {
            $_.ProviderName -eq 'Outlook' -and
            $_.Message -match [regex]::Escape($progId)
        } |
        Select-Object TimeCreated, Id, RecordId, Message)
}

function Get-Median {
    param([int[]]$Values)

    if ($null -eq $Values -or $Values.Count -eq 0) {
        return $null
    }

    $ordered = @($Values | Sort-Object)
    $middle = [int][Math]::Floor($ordered.Count / 2)
    if (($ordered.Count % 2) -eq 1) {
        return [double]$ordered[$middle]
    }

    return ($ordered[$middle - 1] + $ordered[$middle]) / 2.0
}

function Show-Status {
    $state = $null
    $since = $null
    if (Test-Path -LiteralPath $stateFilePath) {
        $state = Get-Content -LiteralPath $stateFilePath -Raw | ConvertFrom-Json
        $since = [datetime]::Parse(
            $state.StartedAt,
            [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::RoundtripKind)
    }

    $loadBehavior = Get-RegistryValueSnapshot -Path $addinRegistryPath -Name 'LoadBehavior'
    $delay = Get-RegistryValueSnapshot -Path $diagnosticsRegistryPath -Name $startupDelayValueName
    $exception = Get-RegistryValueSnapshot -Path $doNotDisableRegistryPath -Name $progId
    $managed = Get-RegistryValueSnapshot -Path $managedAddinsRegistryPath -Name $progId
    $policy = Get-RegistryValueSnapshot -Path $policyAddinsRegistryPath -Name $progId
    $outlookRunning = [bool](Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)

    Write-Host 'Outcom - Outlook add-in resiliency test'
    Write-Host ("Test active             : {0}" -f [bool]$state)
    Write-Host ("Outlook running         : {0}" -f $outlookRunning)
    Write-Host ("LoadBehavior            : {0}" -f $(if ($loadBehavior.Exists) { $loadBehavior.Value } else { '<absent>' }))
    Write-Host ("Debug startup delay (ms): {0}" -f $(if ($delay.Exists) { $delay.Value } else { '<absent>' }))
    Write-Host ("User performance exempt : {0}" -f $(if ($exception.Exists) { $exception.Value } else { '<absent>' }))
    Write-Host ("Managed add-in setting  : {0}" -f $(if ($managed.Exists) { $managed.Value } else { '<absent>' }))
    Write-Host ("Policy add-in setting   : {0}" -f $(if ($policy.Exists) { $policy.Value } else { '<absent>' }))

    $startupEvents = @(Get-OutcomStartupEvents -Since $since)
    if ($null -eq $since) {
        $startupEvents = @($startupEvents | Select-Object -First 5)
        Write-Host 'Recent startup measurements (Event 45):'
    }
    else {
        Write-Host ("Startup measurements since {0:u} (Event 45):" -f $since)
    }

    if ($startupEvents.Count -eq 0) {
        Write-Host '  <none>'
    }
    else {
        foreach ($measurement in $startupEvents) {
            Write-Host ("  {0:u}  {1} ms  record {2}" -f `
                $measurement.TimeCreated,
                $measurement.DurationMilliseconds,
                $measurement.RecordId)
        }

        $latestFive = @($startupEvents | Select-Object -First 5)
        $median = Get-Median -Values @($latestFive.DurationMilliseconds)
        Write-Host ("Median of displayed latest five: {0:N0} ms (threshold: 1000 ms)" -f $median)
    }

    $disableEvents = @(Get-OutcomDisableEvents -Since $since)
    Write-Host ("Disable events for Outcom (Event 59): {0}" -f $disableEvents.Count)
    foreach ($disableEvent in $disableEvents) {
        Write-Host ("  {0:u}  record {1}" -f $disableEvent.TimeCreated, $disableEvent.RecordId)
    }

    if ($state -and $disableEvents.Count -gt 0) {
        Write-Host 'RESULT: Outlook recorded automatic disabling of Outcom.' -ForegroundColor Green
    }
    elseif ($state -and $startupEvents.Count -ge 5) {
        Write-Host 'RESULT: five starts were measured, but no automatic disable event was found.' -ForegroundColor Yellow
    }
}

switch ($Action) {
    'Prepare' {
        Assert-OutlookStopped

        if (Test-Path -LiteralPath $stateFilePath) {
            throw "A test is already active. Run with -Action Restore first: $stateFilePath"
        }

        if (-not (Test-Path -LiteralPath $addinRegistryPath)) {
            throw "The Outlook add-in is not registered: $addinRegistryPath"
        }

        $manifest = Get-RegistryValueSnapshot -Path $addinRegistryPath -Name 'Manifest'
        if (-not $manifest.Exists -or $manifest.Value -notmatch '[\\/]bin[\\/]Debug[\\/]') {
            throw 'The registered add-in is not the Debug build. The artificial delay is unavailable in Release.'
        }

        $managed = Get-RegistryValueSnapshot -Path $managedAddinsRegistryPath -Name $progId
        $policy = Get-RegistryValueSnapshot -Path $policyAddinsRegistryPath -Name $progId
        foreach ($setting in @($managed, $policy)) {
            if ($setting.Exists -and ([int]$setting.Value -eq 0 -or [int]$setting.Value -eq 1)) {
                throw 'A managed add-in setting blocks this test (always disabled or always enabled).'
            }
        }

        $state = [ordered]@{
            StartedAt = [DateTimeOffset]::Now.ToString('o')
            StartupDelayMilliseconds = $StartupDelayMilliseconds
            LoadBehavior = Get-RegistryValueSnapshot -Path $addinRegistryPath -Name 'LoadBehavior'
            StartupDelay = Get-RegistryValueSnapshot -Path $diagnosticsRegistryPath -Name $startupDelayValueName
            DoNotDisable = Get-RegistryValueSnapshot -Path $doNotDisableRegistryPath -Name $progId
        }

        New-Item -ItemType Directory -Path $stateDirectoryPath -Force | Out-Null
        $state | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $stateFilePath -Encoding UTF8

        try {
            New-ItemProperty `
                -LiteralPath $addinRegistryPath `
                -Name 'LoadBehavior' `
                -Value 3 `
                -PropertyType DWord `
                -Force | Out-Null
            New-Item -Path $diagnosticsRegistryPath -Force | Out-Null
            New-ItemProperty `
                -LiteralPath $diagnosticsRegistryPath `
                -Name $startupDelayValueName `
                -Value $StartupDelayMilliseconds `
                -PropertyType DWord `
                -Force | Out-Null
            if (Test-Path -LiteralPath $doNotDisableRegistryPath) {
                Remove-ItemProperty `
                    -LiteralPath $doNotDisableRegistryPath `
                    -Name $progId `
                    -ErrorAction SilentlyContinue
            }
        }
        catch {
            Set-RegistryValueFromSnapshot -Path $addinRegistryPath -Name 'LoadBehavior' -Snapshot $state.LoadBehavior
            Set-RegistryValueFromSnapshot -Path $diagnosticsRegistryPath -Name $startupDelayValueName -Snapshot $state.StartupDelay
            Set-RegistryValueFromSnapshot -Path $doNotDisableRegistryPath -Name $progId -Snapshot $state.DoNotDisable
            Remove-Item -LiteralPath $stateFilePath -Force -ErrorAction SilentlyContinue
            throw
        }

        Write-Host "Prepared a $StartupDelayMilliseconds ms startup delay for the Debug build." -ForegroundColor Green
        Write-Host 'Start and close Outlook normally five times, then run:'
        Write-Host '  .\tools\Test-OutlookAddinResiliency.ps1 -Action Status'
        Write-Host 'When the test is complete, close Outlook and always run:'
        Write-Host '  .\tools\Test-OutlookAddinResiliency.ps1 -Action Restore'
    }

    'Status' {
        Show-Status
    }

    'Restore' {
        Assert-OutlookStopped

        if (-not (Test-Path -LiteralPath $stateFilePath)) {
            throw "No active test state was found: $stateFilePath"
        }

        $state = Get-Content -LiteralPath $stateFilePath -Raw | ConvertFrom-Json
        Set-RegistryValueFromSnapshot -Path $addinRegistryPath -Name 'LoadBehavior' -Snapshot $state.LoadBehavior
        Set-RegistryValueFromSnapshot -Path $diagnosticsRegistryPath -Name $startupDelayValueName -Snapshot $state.StartupDelay
        Set-RegistryValueFromSnapshot -Path $doNotDisableRegistryPath -Name $progId -Snapshot $state.DoNotDisable
        Remove-Item -LiteralPath $stateFilePath -Force

        Write-Host 'The pre-test Outcom registry values were restored.' -ForegroundColor Green
        Write-Host 'Restart Outlook and verify that the Outcom ribbon is visible.'
    }
}
