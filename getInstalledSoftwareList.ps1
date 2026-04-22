<#
.SYNOPSIS
    Exports installed software to a CSV file.

.DESCRIPTION
    Reads installed software from the standard Windows uninstall registry locations,
    removes obvious non-user-facing components by default, deduplicates entries by
    software name, and writes the result to a CSV file. If winget is available,
    the script will also try to populate a newer available version for matched apps.

.PARAMETER OutputPath
    Optional path for the CSV file. Defaults to InstalledSoftware.csv beside this script.

.PARAMETER IncludeSystemComponents
    Includes updates, drivers, and other component-style entries that are filtered
    out by default.

.EXAMPLE
    .\getInstalledSoftwareList.ps1
    Writes InstalledSoftware.csv beside the script.

.EXAMPLE
    .\getInstalledSoftwareList.ps1 -OutputPath C:\Temp\InstalledSoftware.csv -IncludeSystemComponents
    Writes the full inventory, including system components, to C:\Temp\InstalledSoftware.csv.
#>

[CmdletBinding()]
param(
    [string]$OutputPath,

    [switch]$IncludeSystemComponents
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $scriptDirectory = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Path $MyInvocation.MyCommand.Path -Parent }
    $OutputPath = Join-Path -Path $scriptDirectory -ChildPath 'InstalledSoftware.csv'
}

function Get-NormalizedDisplayName {
    param(
        [AllowNull()]
        [string]$Name
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return $null
    }

    $normalized = $Name.Trim()
    $normalized = $normalized -replace '[\u2122\u00AE\u00A9]', ''
    $normalized = $normalized -replace '\s+', ' '
    return $normalized.Trim()
}

function Get-MatchKey {
    param(
        [AllowNull()]
        [string]$Name
    )

    $normalized = Get-NormalizedDisplayName -Name $Name
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return $null
    }

    $key = $normalized.ToLowerInvariant()
    $key = $key -replace '\(x(86|64)\)', ' '
    $key = $key -replace '\b(32-bit|64-bit|x86|x64)\b', ' '
    $key = $key -replace '[^a-z0-9]+', ' '
    $key = $key -replace '\b(edition|runtime|redistributable|redistributables|driver|drivers)\b', ' '
    $key = $key -replace '\s+', ' '
    return $key.Trim()
}

function Get-VersionObject {
    param(
        [AllowNull()]
        [string]$Version
    )

    if ([string]::IsNullOrWhiteSpace($Version)) {
        return $null
    }

    $trimmed = $Version.Trim()
    $candidate = $trimmed -replace '[^0-9\.]', '.'
    $candidate = $candidate -replace '\.{2,}', '.'
    $candidate = $candidate.Trim('.')

    if ([string]::IsNullOrWhiteSpace($candidate)) {
        return $null
    }

    try {
        return [version]$candidate
    }
    catch {
        return $null
    }
}

function Compare-VersionStrings {
    param(
        [AllowNull()]
        [string]$Left,

        [AllowNull()]
        [string]$Right
    )

    $leftVersion = Get-VersionObject -Version $Left
    $rightVersion = Get-VersionObject -Version $Right

    if ($leftVersion -and $rightVersion) {
        return $leftVersion.CompareTo($rightVersion)
    }

    if (-not [string]::IsNullOrWhiteSpace($Left) -and [string]::IsNullOrWhiteSpace($Right)) {
        return 1
    }

    if ([string]::IsNullOrWhiteSpace($Left) -and -not [string]::IsNullOrWhiteSpace($Right)) {
        return -1
    }

    return [string]::Compare($Left, $Right, $true)
}

function Test-IsLikelyComponentName {
    param(
        [string]$Name
    )

    $patterns = @(
        'security update',
        'update for microsoft',
        '^update for ',
        '^hotfix',
        '^kb\d+',
        'cumulative update',
        'servicing stack',
        'redistributable',
        '\bruntime\b',
        '\bdriver\b',
        '\bdrivers\b',
        'firmware',
        'language pack',
        'webview2 runtime'
    )

    foreach ($pattern in $patterns) {
        if ($Name -match $pattern) {
            return $true
        }
    }

    return $false
}

function Get-PropertyValue {
    param(
        [psobject]$Object,
        [string]$Name
    )

    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}

function Test-IsUserFacingSoftware {
    param(
        [psobject]$Entry
    )

    if ($IncludeSystemComponents.IsPresent) {
        return $true
    }

    $displayName = [string](Get-PropertyValue -Object $Entry -Name 'DisplayName')
    if ([string]::IsNullOrWhiteSpace($displayName)) {
        return $false
    }

    $systemComponent = Get-PropertyValue -Object $Entry -Name 'SystemComponent'
    if ($systemComponent -eq 1) {
        return $false
    }

    $releaseType = [string](Get-PropertyValue -Object $Entry -Name 'ReleaseType')
    if ($releaseType -match 'Hotfix|Security Update|Update Rollup|ServicePack') {
        return $false
    }

    $parentKeyName = [string](Get-PropertyValue -Object $Entry -Name 'ParentKeyName')
    if (-not [string]::IsNullOrWhiteSpace($parentKeyName)) {
        return $false
    }

    $name = $displayName.Trim()
    $publisher = [string](Get-PropertyValue -Object $Entry -Name 'Publisher')

    if (Test-IsLikelyComponentName -Name $name.ToLowerInvariant()) {
        $allowPatterns = @(
            '^microsoft visual studio',
            '^visual studio code$',
            '^microsoft 365',
            '^microsoft office',
            '^python ',
            '^node\.js',
            '^git$',
            '^docker desktop',
            '^google chrome$',
            '^mozilla firefox$'
        )

        foreach ($pattern in $allowPatterns) {
            if ($name.ToLowerInvariant() -match $pattern) {
                return $true
            }
        }

        return $false
    }

    if ($publisher -match 'Microsoft Corporation' -and $name -match '^Microsoft (Windows|Edge WebView2 Runtime|Visual C\+\+ 20\d{2} Redistributable)') {
        return $false
    }

    return $true
}

function Get-RegistrySoftwareEntries {
    $registryPaths = @(
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    foreach ($path in $registryPaths) {
        Write-Host "Scanning $path" -ForegroundColor Cyan

        if (-not (Test-Path -Path $path)) {
            continue
        }

        Get-ItemProperty -Path $path -ErrorAction SilentlyContinue |
            Where-Object { -not [string]::IsNullOrWhiteSpace([string](Get-PropertyValue -Object $_ -Name 'DisplayName')) } |
            ForEach-Object {
                if (-not (Test-IsUserFacingSoftware -Entry $_)) {
                    return
                }

                $normalizedName = Get-NormalizedDisplayName -Name ([string](Get-PropertyValue -Object $_ -Name 'DisplayName'))
                if ([string]::IsNullOrWhiteSpace($normalizedName)) {
                    return
                }

                [pscustomobject]@{
                    Softwarename      = $normalizedName
                    'Current Version' = ([string](Get-PropertyValue -Object $_ -Name 'DisplayVersion')).Trim()
                    'Latest Version'  = ''
                    MatchKey          = Get-MatchKey -Name $normalizedName
                    Source            = 'Registry'
                }
            }
    }
}

function Select-BestSoftwareEntry {
    param(
        [object[]]$Entries
    )

    $selected = $Entries | Select-Object -First 1

    foreach ($candidate in ($Entries | Select-Object -Skip 1)) {
        $currentVersionComparison = Compare-VersionStrings -Left $candidate.'Current Version' -Right $selected.'Current Version'
        if ($currentVersionComparison -gt 0) {
            $selected = $candidate
            continue
        }

        if ($currentVersionComparison -eq 0) {
            $candidateNameLength = ($candidate.Softwarename | Measure-Object -Character).Characters
            $selectedNameLength = ($selected.Softwarename | Measure-Object -Character).Characters
            if ($candidateNameLength -gt $selectedNameLength) {
                $selected = $candidate
            }
        }
    }

    return $selected
}

function ConvertFrom-WingetTable {
    param(
        [string[]]$Lines
    )

    $headerIndex = -1
    $separatorIndex = -1

    for ($i = 0; $i -lt $Lines.Count - 1; $i++) {
        if ($Lines[$i] -match '^\s*Name\s+Id\s+') {
            $headerIndex = $i
            $separatorIndex = $i + 1
            break
        }
    }

    if ($headerIndex -lt 0 -or $separatorIndex -ge $Lines.Count) {
        return @()
    }

    $headerLine = $Lines[$headerIndex]
    $separatorLine = $Lines[$separatorIndex]

    $nameStart = $headerLine.IndexOf('Name')
    $idStart = $headerLine.IndexOf('Id', $nameStart)
    $versionStart = $headerLine.IndexOf('Version', $idStart)
    $availableStart = $headerLine.IndexOf('Available', $versionStart)
    $sourceStart = $headerLine.IndexOf('Source', [Math]::Max($availableStart, $versionStart))

    if ($nameStart -lt 0 -or $idStart -lt 0 -or $versionStart -lt 0) {
        return @()
    }

    $dataStart = $separatorIndex + 1
    $results = New-Object System.Collections.Generic.List[object]

    for ($i = $dataStart; $i -lt $Lines.Count; $i++) {
        $line = $Lines[$i]
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        if ($line -match '^\s*[-]+$') {
            continue
        }

        if ($line -match '^\s*\d+\s+upgrades available') {
            continue
        }

        $name = if ($idStart -gt $nameStart) { $line.Substring($nameStart, [Math]::Min($idStart - $nameStart, [Math]::Max($line.Length - $nameStart, 0))).Trim() } else { '' }
        $id = if ($versionStart -gt $idStart -and $line.Length -gt $idStart) { $line.Substring($idStart, [Math]::Min($versionStart - $idStart, $line.Length - $idStart)).Trim() } else { '' }

        $versionWidth = if ($availableStart -gt $versionStart) { $availableStart - $versionStart } elseif ($sourceStart -gt $versionStart) { $sourceStart - $versionStart } else { $line.Length - $versionStart }
        $version = if ($line.Length -gt $versionStart) { $line.Substring($versionStart, [Math]::Min($versionWidth, $line.Length - $versionStart)).Trim() } else { '' }

        $available = ''
        if ($availableStart -gt 0 -and $line.Length -gt $availableStart) {
            $availableWidth = if ($sourceStart -gt $availableStart) { $sourceStart - $availableStart } else { $line.Length - $availableStart }
            $available = $line.Substring($availableStart, [Math]::Min($availableWidth, $line.Length - $availableStart)).Trim()
        }

        $source = if ($sourceStart -gt 0 -and $line.Length -gt $sourceStart) { $line.Substring($sourceStart).Trim() } else { '' }

        if ([string]::IsNullOrWhiteSpace($name) -or [string]::IsNullOrWhiteSpace($id)) {
            continue
        }

        $results.Add([pscustomobject]@{
            Name      = $name
            Id        = $id
            Version   = $version
            Available = $available
            Source    = $source
            MatchKey  = Get-MatchKey -Name $name
        })
    }

    return $results
}

function Get-WingetCommand {
    $candidate = Get-Command -Name 'winget.exe' -ErrorAction SilentlyContinue
    if ($candidate) {
        return $candidate.Source
    }

    return $null
}

function Invoke-WingetText {
    param(
        [string[]]$Arguments
    )

    $output = & $script:WingetPath @Arguments 2>&1
    return @($output | ForEach-Object { "$_" })
}

function Get-WingetInstalledPackages {
    Write-Host 'Collecting installed package names from winget...' -ForegroundColor Cyan
    $lines = Invoke-WingetText -Arguments @(
        'list',
        '--accept-source-agreements',
        '--disable-interactivity'
    )

    return ConvertFrom-WingetTable -Lines $lines
}

function Get-WingetUpgrades {
    Write-Host 'Collecting available upgrades from winget...' -ForegroundColor Cyan
    $lines = Invoke-WingetText -Arguments @(
        'upgrade',
        '--accept-source-agreements',
        '--disable-interactivity',
        '--include-unknown'
    )

    return ConvertFrom-WingetTable -Lines $lines
}

function Test-CredibleWingetMatch {
    param(
        [psobject]$RegistryEntry,
        [psobject]$WingetInstalledEntry,
        [psobject]$WingetUpgradeEntry
    )

    if (-not $WingetInstalledEntry -or -not $WingetUpgradeEntry) {
        return $false
    }

    if ([string]::IsNullOrWhiteSpace($RegistryEntry.MatchKey) -or [string]::IsNullOrWhiteSpace($WingetInstalledEntry.MatchKey)) {
        return $false
    }

    if ($RegistryEntry.MatchKey -ne $WingetInstalledEntry.MatchKey) {
        return $false
    }

    if ($WingetInstalledEntry.Id -ne $WingetUpgradeEntry.Id) {
        return $false
    }

    $currentVersion = $RegistryEntry.'Current Version'
    $wingetCurrent = $WingetInstalledEntry.Version
    $wingetAvailable = $WingetUpgradeEntry.Available

    if ([string]::IsNullOrWhiteSpace($wingetAvailable)) {
        return $false
    }

    if (-not [string]::IsNullOrWhiteSpace($currentVersion) -and -not [string]::IsNullOrWhiteSpace($wingetCurrent)) {
        $versionComparison = Compare-VersionStrings -Left $currentVersion -Right $wingetCurrent
        if ([Math]::Abs($versionComparison) -gt 0 -and $RegistryEntry.Softwarename -ne $WingetInstalledEntry.Name) {
            return $false
        }
    }

    return (Compare-VersionStrings -Left $wingetAvailable -Right $currentVersion) -gt 0
}

function Add-WingetLatestVersions {
    param(
        [System.Collections.Generic.List[object]]$Software
    )

    $script:WingetPath = Get-WingetCommand
    if (-not $script:WingetPath) {
        Write-Host 'winget not found. Latest Version values will remain blank.' -ForegroundColor Yellow
        return
    }

    Write-Host "Using winget at $script:WingetPath" -ForegroundColor Cyan

    try {
        $wingetInstalled = Get-WingetInstalledPackages
        $wingetUpgrades = Get-WingetUpgrades
    }
    catch {
        Write-Warning "winget queries failed: $($_.Exception.Message)"
        return
    }

    if (-not $wingetInstalled -or -not $wingetUpgrades) {
        return
    }

    $installedByKey = @{}
    foreach ($item in $wingetInstalled) {
        if ([string]::IsNullOrWhiteSpace($item.MatchKey)) {
            continue
        }

        if (-not $installedByKey.ContainsKey($item.MatchKey)) {
            $installedByKey[$item.MatchKey] = $item
        }
    }

    $upgradesByKey = @{}
    foreach ($item in $wingetUpgrades) {
        if ([string]::IsNullOrWhiteSpace($item.MatchKey)) {
            continue
        }

        if (-not $upgradesByKey.ContainsKey($item.MatchKey)) {
            $upgradesByKey[$item.MatchKey] = $item
        }
    }

    foreach ($entry in $Software) {
        if ([string]::IsNullOrWhiteSpace($entry.MatchKey)) {
            continue
        }

        $wingetInstalledEntry = $installedByKey[$entry.MatchKey]
        $wingetUpgradeEntry = $upgradesByKey[$entry.MatchKey]

        if (Test-CredibleWingetMatch -RegistryEntry $entry -WingetInstalledEntry $wingetInstalledEntry -WingetUpgradeEntry $wingetUpgradeEntry) {
            $entry.'Latest Version' = $wingetUpgradeEntry.Available
        }
    }
}

Write-Host 'Discovering installed software from the registry...' -ForegroundColor Green
$registryEntries = @(Get-RegistrySoftwareEntries)

if (-not $registryEntries) {
    Write-Warning 'No installed software entries were found.'
}

Write-Host 'Deduplicating software entries...' -ForegroundColor Green
$deduplicated = New-Object System.Collections.Generic.List[object]

foreach ($group in ($registryEntries | Group-Object -Property Softwarename | Sort-Object -Property Name)) {
    $selected = Select-BestSoftwareEntry -Entries $group.Group
    $deduplicated.Add($selected)
}

Add-WingetLatestVersions -Software $deduplicated

$finalOutput = $deduplicated |
    Sort-Object -Property Softwarename |
    Select-Object -Property Softwarename, 'Current Version', 'Latest Version'

$outputDirectory = Split-Path -Path $OutputPath -Parent
if (-not [string]::IsNullOrWhiteSpace($outputDirectory) -and -not (Test-Path -Path $outputDirectory)) {
    New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
}

Write-Host "Writing CSV to $OutputPath" -ForegroundColor Green
$finalOutput | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8

Write-Host "Completed. Exported $($finalOutput.Count) software entries." -ForegroundColor Green
