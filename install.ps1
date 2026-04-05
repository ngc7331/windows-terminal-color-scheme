[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$BuiltInSchemes = @(
    'CGA',
    'Campbell',
    'Campbell Powershell',
    'Dark+',
    'Dimidium',
    'IBM 5153',
    'One Half Dark',
    'One Half Light',
    'Ottosson',
    'Solarized Dark',
    'Solarized Light',
    'Tango Dark',
    'Tango Light',
    'Vintage'
)

function Get-WtSettingsCandidates {
    $localAppData = $null
    $candidateRoots = @(
        $env:LOCALAPPDATA,
        [Environment]::GetEnvironmentVariable('LOCALAPPDATA', 'User'),
        [Environment]::GetEnvironmentVariable('LOCALAPPDATA', 'Process'),
        [Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)
    )

    foreach ($root in $candidateRoots) {
        if (-not [string]::IsNullOrWhiteSpace($root)) {
            $localAppData = $root
            break
        }
    }

    if ([string]::IsNullOrWhiteSpace($localAppData) -and -not [string]::IsNullOrWhiteSpace($env:USERPROFILE)) {
        $localAppData = Join-Path $env:USERPROFILE 'AppData/Local'
    }

    if ([string]::IsNullOrWhiteSpace($localAppData)) {
        return @()
    }

    return @(
        [PSCustomObject]@{
            Label = 'Windows Terminal (Stable)'
            Path  = Join-Path $localAppData 'Packages/Microsoft.WindowsTerminal_8wekyb3d8bbwe/LocalState/settings.json'
        },
        [PSCustomObject]@{
            Label = 'Windows Terminal (Preview)'
            Path  = Join-Path $localAppData 'Packages/Microsoft.WindowsTerminalPreview_8wekyb3d8bbwe/LocalState/settings.json'
        },
        [PSCustomObject]@{
            Label = 'Windows Terminal (Unpackaged: Scoop/Chocolatey)'
            Path  = Join-Path $localAppData 'Microsoft/Windows Terminal/settings.json'
        }
    )
}

function Select-WtSettingsPath {
    $candidates = Get-WtSettingsCandidates
    $existing = @($candidates | Where-Object { Test-Path -LiteralPath $_.Path })

    if ($existing.Count -eq 0) {
        if (@($candidates).Count -eq 0) {
            Write-Host 'Cannot auto-detect LOCALAPPDATA in current session.' -ForegroundColor Yellow
        }
        else {
            Write-Host 'No settings.json found in known locations.' -ForegroundColor Yellow
        }
        Write-Host 'Enter your settings.json path manually:' -ForegroundColor Yellow
        $manualPath = Read-Host 'settings.json path'
        if ([string]::IsNullOrWhiteSpace($manualPath) -or -not (Test-Path -LiteralPath $manualPath)) {
            throw 'Invalid settings.json path.'
        }
        return (Resolve-Path -LiteralPath $manualPath).Path
    }

    if ($existing.Count -eq 1) {
        return $existing[0].Path
    }

    Write-Host 'Multiple settings.json files found. Choose one:' -ForegroundColor Cyan
    for ($i = 0; $i -lt $existing.Count; $i++) {
        Write-Host ("{0}. {1}: {2}" -f ($i + 1), $existing[$i].Label, $existing[$i].Path)
    }

    $selection = Read-Host 'Enter number'
    $index = 0
    if (-not [int]::TryParse($selection, [ref]$index)) {
        throw 'Selection is not a valid number.'
    }

    if ($index -lt 1 -or $index -gt $existing.Count) {
        throw 'Selection out of range.'
    }

    return $existing[$index - 1].Path
}

function Read-JsonObject {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($raw)) {
        throw "JSON file is empty: $Path"
    }

    return $raw | ConvertFrom-Json
}

function Write-JsonObject {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [psobject]$Object
    )

    $json = $Object | ConvertTo-Json -Depth 100
    Set-Content -LiteralPath $Path -Value $json -Encoding UTF8
}

function Get-RepoSchemes {
    $schemeDir = Join-Path $PSScriptRoot 'scheme'
    if (-not (Test-Path -LiteralPath $schemeDir)) {
        throw "Scheme directory not found: $schemeDir"
    }

    $files = Get-ChildItem -LiteralPath $schemeDir -Filter '*.json' -File | Sort-Object Name
    if ($files.Count -eq 0) {
        throw "No scheme json files found in: $schemeDir"
    }

    $schemes = @()
    foreach ($file in $files) {
        $obj = Read-JsonObject -Path $file.FullName
        if (-not $obj.name) {
            throw "Missing 'name' in scheme file: $($file.FullName)"
        }
        $schemes += $obj
    }

    return $schemes
}

function Ensure-SchemesArray {
    param([psobject]$Settings)

    if (-not $Settings.PSObject.Properties.Match('schemes')) {
        $Settings | Add-Member -MemberType NoteProperty -Name schemes -Value @()
        return
    }

    if ($null -eq $Settings.schemes) {
        $Settings.schemes = @()
    }
}

function Set-OrAddProperty {
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Object,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [object]$Value
    )

    if ($Object.PSObject.Properties.Match($Name)) {
        $Object.$Name = $Value
        return
    }

    $Object | Add-Member -MemberType NoteProperty -Name $Name -Value $Value
}

function Ensure-ProfilesDefaults {
    param([psobject]$Settings)

    if (-not $Settings.PSObject.Properties.Match('profiles')) {
        $Settings | Add-Member -MemberType NoteProperty -Name profiles -Value ([PSCustomObject]@{})
    }

    if ($null -eq $Settings.profiles) {
        $Settings.profiles = [PSCustomObject]@{}
    }

    if (-not $Settings.profiles.PSObject.Properties.Match('defaults')) {
        $Settings.profiles | Add-Member -MemberType NoteProperty -Name defaults -Value ([PSCustomObject]@{})
    }

    if ($null -eq $Settings.profiles.defaults) {
        $Settings.profiles.defaults = [PSCustomObject]@{}
    }
}

function Ensure-DefaultModeColorSchemeObject {
    param([psobject]$Settings)

    Ensure-ProfilesDefaults -Settings $Settings

    if (-not $Settings.profiles.defaults.PSObject.Properties.Match('colorScheme')) {
        $Settings.profiles.defaults | Add-Member -MemberType NoteProperty -Name colorScheme -Value ([PSCustomObject]@{ dark = 'Campbell'; light = 'Campbell' })
        return
    }

    $existing = $Settings.profiles.defaults.colorScheme
    if ($existing -is [string]) {
        $Settings.profiles.defaults.colorScheme = [PSCustomObject]@{
            dark  = $existing
            light = $existing
        }
        return
    }

    if ($null -eq $existing) {
        $Settings.profiles.defaults.colorScheme = [PSCustomObject]@{ dark = 'Campbell'; light = 'Campbell' }
        return
    }

    $darkValue = 'Campbell'
    $lightValue = 'Campbell'

    if ($existing.PSObject.Properties.Match('dark') -and -not [string]::IsNullOrWhiteSpace($existing.dark)) {
        $darkValue = $existing.dark
    }
    if ($existing.PSObject.Properties.Match('light') -and -not [string]::IsNullOrWhiteSpace($existing.light)) {
        $lightValue = $existing.light
    }

    $Settings.profiles.defaults.colorScheme = [PSCustomObject]@{
        dark  = $darkValue
        light = $lightValue
    }
}

function Get-SelectableSchemeNames {
    param(
        [psobject]$Settings,
        [array]$RepoSchemes
    )

    $ordered = [System.Collections.Generic.List[string]]::new()
    foreach ($name in $BuiltInSchemes) {
        if (-not [string]::IsNullOrWhiteSpace($name)) {
            [void]$ordered.Add($name)
        }
    }

    Ensure-SchemesArray -Settings $Settings
    foreach ($item in $Settings.schemes) {
        if ($null -ne $item -and -not [string]::IsNullOrWhiteSpace($item.name)) {
            [void]$ordered.Add($item.name)
        }
    }

    foreach ($item in $RepoSchemes) {
        if ($null -ne $item -and -not [string]::IsNullOrWhiteSpace($item.name)) {
            [void]$ordered.Add($item.name)
        }
    }

    $seen = @{}
    $result = @()
    foreach ($name in $ordered) {
        if (-not $seen.ContainsKey($name)) {
            $seen[$name] = $true
            $result += $name
        }
    }

    return $result
}

function Select-SchemeName {
    param(
        [array]$SelectableSchemeNames,
        [string]$Mode
    )

    Write-Host ''
    Write-Host ("Available schemes for {0} mode:" -f $Mode) -ForegroundColor Cyan
    for ($i = 0; $i -lt $SelectableSchemeNames.Count; $i++) {
        Write-Host ("{0}. {1}" -f ($i + 1), $SelectableSchemeNames[$i])
    }

    $selection = Read-Host 'Select scheme number'
    $index = 0
    if (-not [int]::TryParse($selection, [ref]$index)) {
        throw 'Selection is not a valid number.'
    }

    if ($index -lt 1 -or $index -gt $SelectableSchemeNames.Count) {
        throw 'Selection out of range.'
    }

    return $SelectableSchemeNames[$index - 1]
}

function Ensure-SchemeExistsIfFromRepo {
    param(
        [psobject]$Settings,
        [array]$RepoSchemes,
        [string]$SelectedSchemeName
    )

    $repoScheme = $RepoSchemes | Where-Object { $_.name -eq $SelectedSchemeName } | Select-Object -First 1
    if ($null -eq $repoScheme) {
        return
    }

    Ensure-SchemesArray -Settings $Settings
    $exists = @($Settings.schemes | Where-Object { $_.name -eq $SelectedSchemeName }).Count -gt 0
    if (-not $exists) {
        $Settings.schemes = @($Settings.schemes + $repoScheme)
    }
}

function Set-ModeColorScheme {
    param(
        [psobject]$Settings,
        [array]$RepoSchemes,
        [ValidateSet('dark', 'light')]
        [string]$Mode
    )

    $selectable = Get-SelectableSchemeNames -Settings $Settings -RepoSchemes $RepoSchemes
    if ($selectable.Count -eq 0) {
        throw 'No schemes available to select.'
    }

    $selectedName = Select-SchemeName -SelectableSchemeNames $selectable -Mode $Mode
    Ensure-SchemeExistsIfFromRepo -Settings $Settings -RepoSchemes $RepoSchemes -SelectedSchemeName $selectedName
    Ensure-DefaultModeColorSchemeObject -Settings $Settings
    Set-OrAddProperty -Object $Settings.profiles.defaults.colorScheme -Name $Mode -Value $selectedName

    Write-Host ("Set profiles.defaults.colorScheme.{0} = '{1}'." -f $Mode, $selectedName) -ForegroundColor Green
}

function Reset-ConfiguredRepoSchemesToCampbell {
    param(
        [psobject]$Settings,
        [array]$RepoSchemes
    )

    $repoNameSet = @{}
    foreach ($item in $RepoSchemes) {
        if ($null -ne $item -and -not [string]::IsNullOrWhiteSpace($item.name)) {
            $repoNameSet[$item.name] = $true
        }
    }

    $resetCount = 0

    $hasProperty = {
        param(
            [object]$obj,
            [string]$name
        )

        if ($null -eq $obj) {
            return $false
        }

        return $obj.PSObject.Properties.Match($name).Count -gt 0
    }

    $resetColorSchemeValue = {
        param(
            [object]$container,
            [ref]$counter
        )

        if (-not (& $hasProperty $container 'colorScheme')) {
            return
        }

        $value = $container.colorScheme
        if ($value -is [string]) {
            if ($repoNameSet.ContainsKey($value)) {
                $container.colorScheme = 'Campbell'
                $counter.Value += 1
            }
            return
        }

        if ($null -eq $value) {
            return
        }

        if (($value.PSObject.Properties.Match('dark').Count -gt 0) -and $repoNameSet.ContainsKey($value.dark)) {
            $value.dark = 'Campbell'
            $counter.Value += 1
        }
        if (($value.PSObject.Properties.Match('light').Count -gt 0) -and $repoNameSet.ContainsKey($value.light)) {
            $value.light = 'Campbell'
            $counter.Value += 1
        }
    }

    if ((& $hasProperty $Settings 'profiles') -and $null -ne $Settings.profiles) {
        if ((& $hasProperty $Settings.profiles 'defaults') -and $null -ne $Settings.profiles.defaults) {
            & $resetColorSchemeValue $Settings.profiles.defaults ([ref]$resetCount)
        }

        if ((& $hasProperty $Settings.profiles 'list') -and $null -ne $Settings.profiles.list) {
            foreach ($profile in $Settings.profiles.list) {
                & $resetColorSchemeValue $profile ([ref]$resetCount)
            }
        }
    }

    return $resetCount
}

function Install-RepoSchemes {
    param(
        [psobject]$Settings,
        [array]$RepoSchemes
    )

    Ensure-SchemesArray -Settings $Settings
    $repoNames = @($RepoSchemes | ForEach-Object { $_.name })

    $kept = @($Settings.schemes | Where-Object { $repoNames -notcontains $_.name })
    $Settings.schemes = @($kept + $RepoSchemes)

    Write-Host ("Installed/Updated {0} schemes." -f $RepoSchemes.Count) -ForegroundColor Green
}

function Uninstall-RepoSchemes {
    param(
        [psobject]$Settings,
        [array]$RepoSchemes
    )

    Ensure-SchemesArray -Settings $Settings
    $repoNames = @($RepoSchemes | ForEach-Object { $_.name })

    $before = @($Settings.schemes).Count
    $Settings.schemes = @($Settings.schemes | Where-Object { $repoNames -notcontains $_.name })
    $after = @($Settings.schemes).Count
    $resetCount = Reset-ConfiguredRepoSchemesToCampbell -Settings $Settings -RepoSchemes $RepoSchemes

    Write-Host ("Removed {0} schemes." -f ($before - $after)) -ForegroundColor Green
    if ($resetCount -gt 0) {
        Write-Host ("Reset {0} configured scheme reference(s) to 'Campbell'." -f $resetCount) -ForegroundColor Green
    }
}

function Show-Menu {
    Write-Host ''
    Write-Host 'Windows Terminal Color Scheme Manager' -ForegroundColor Cyan
    Write-Host '1. Install schemes'
    Write-Host '2. Uninstall schemes'
    Write-Host '3. Set dark mode scheme'
    Write-Host '4. Set light mode scheme'
    Write-Host '0. Exit'
    Write-Host ''
}

try {
    $settingsPath = Select-WtSettingsPath
    $repoSchemes = Get-RepoSchemes

    Write-Host ("Using settings file: {0}" -f $settingsPath) -ForegroundColor Cyan

    while ($true) {
        Show-Menu
        $choice = Read-Host 'Select an option'

        if ($choice -eq '0') {
            Write-Host 'Bye.' -ForegroundColor Cyan
            break
        }

        $settings = Read-JsonObject -Path $settingsPath

        switch ($choice) {
            '1' {
                Install-RepoSchemes -Settings $settings -RepoSchemes $repoSchemes
                Write-JsonObject -Path $settingsPath -Object $settings
            }
            '2' {
                Uninstall-RepoSchemes -Settings $settings -RepoSchemes $repoSchemes
                Write-JsonObject -Path $settingsPath -Object $settings
            }
            '3' {
                Set-ModeColorScheme -Settings $settings -RepoSchemes $repoSchemes -Mode 'dark'
                Write-JsonObject -Path $settingsPath -Object $settings
            }
            '4' {
                Set-ModeColorScheme -Settings $settings -RepoSchemes $repoSchemes -Mode 'light'
                Write-JsonObject -Path $settingsPath -Object $settings
            }
            default {
                Write-Host 'Invalid choice, please retry.' -ForegroundColor Yellow
                continue
            }
        }

        Write-Host 'settings.json updated.' -ForegroundColor Green
    }
}
catch {
    Write-Error $_
    exit 1
}
