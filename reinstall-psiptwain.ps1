#Requires -RunAsAdministrator
<#
Author : Joshua Dwight
Github : https://github.com/joshdwight101

Purpose:
- Uninstall and reinstall PaperStream IP TWAIN
- Uninstall and reinstall Software Operation Panel
- Uninstall and reinstall SP Series Online Update
- Uses MSI packages directly instead of the GUI Setup.exe wrapper

Package layout supported:
.\Reinstall-PaperStream-SP.ps1
.\PSIP_SP_TWAIN.msi
.\SOPSetup.msi
.\OLUSetup.msi

Also supported:
- PSIP_TWAIN.msi instead of PSIP_SP_TWAIN.msi
- MSI files inside subfolders if -SearchSubfolders is used
#>

[CmdletBinding()]
param(
    [switch]$SkipUninstall,
    [switch]$SkipInstall,
    [switch]$SkipSOP,
    [switch]$SkipOLU,
    [switch]$SearchSubfolders
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$LogDir    = Join-Path $ScriptDir 'Logs'

$MainLog          = Join-Path $LogDir 'Reinstall-PaperStream-SP.log'
$TwainInstallLog  = Join-Path $LogDir 'Install-PSIP_SP_TWAIN.log'
$SopInstallLog    = Join-Path $LogDir 'Install-SOPSetup.log'
$OluInstallLog    = Join-Path $LogDir 'Install-OLUSetup.log'

New-Item -Path $LogDir -ItemType Directory -Force | Out-Null

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet('INFO','WARN','ERROR')]
        [string]$Level = 'INFO'
    )

    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line
    Add-Content -Path $MainLog -Value $line
}

function Find-MsiByNames {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Names,

        [Parameter(Mandatory = $true)]
        [string]$BasePath,

        [switch]$Recurse
    )

    foreach ($name in $Names) {
        $directPath = Join-Path $BasePath $name
        if (Test-Path -LiteralPath $directPath -PathType Leaf) {
            return (Get-Item -LiteralPath $directPath).FullName
        }
    }

    if ($Recurse) {
        $files = Get-ChildItem -LiteralPath $BasePath -File -Recurse -ErrorAction SilentlyContinue
        foreach ($name in $Names) {
            $match = $files | Where-Object { $_.Name -ieq $name } | Select-Object -First 1
            if ($null -ne $match) {
                return $match.FullName
            }
        }
    }

    return $null
}

$TwainMsi = Find-MsiByNames -Names @('PSIP_SP_TWAIN.msi', 'PSIP_TWAIN.msi') -BasePath $ScriptDir -Recurse:$SearchSubfolders
$SopMsi   = Find-MsiByNames -Names @('SOPSetup.msi') -BasePath $ScriptDir -Recurse:$SearchSubfolders
$OluMsi   = Find-MsiByNames -Names @('OLUSetup.msi') -BasePath $ScriptDir -Recurse:$SearchSubfolders

function Stop-PaperStreamProcesses {
    $names = @(
        'PaperStreamCapture',
        'PaperStreamClickScan',
        'SoftwareOperationPanel',
        'fiScannerAdminTool',
        'fiScannerSetting',
        'TWUNK_32',
        'TWUNK_64'
    )

    foreach ($name in $names) {
        Get-Process -Name $name -ErrorAction SilentlyContinue | ForEach-Object {
            try {
                Write-Log "Stopping process: $($_.ProcessName) (PID $($_.Id))"
                Stop-Process -Id $_.Id -Force -ErrorAction Stop
            }
            catch {
                Write-Log "Could not stop process $($_.ProcessName): $($_.Exception.Message)" 'WARN'
            }
        }
    }
}

function Get-UninstallEntries {
    $roots = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKCU:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    )

    $results = [System.Collections.Generic.List[object]]::new()

    foreach ($root in $roots) {
        if (-not (Test-Path -LiteralPath $root)) {
            continue
        }

        foreach ($subKey in (Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue)) {
            $item = $null

            try {
                $item = Get-ItemProperty -LiteralPath $subKey.PSPath -ErrorAction SilentlyContinue
            }
            catch {
                continue
            }

            if ($null -eq $item) {
                continue
            }

            $propNames = @($item.PSObject.Properties.Name)
            if ('DisplayName' -notin $propNames) {
                continue
            }

            $displayName = [string]$item.DisplayName
            if ([string]::IsNullOrWhiteSpace($displayName)) {
                continue
            }

            $category = $null

            if (
                $displayName -match '^PaperStream IP\s*\(TWAIN\)$' -or
                $displayName -match '^PaperStream IP\s*\(TWAIN x64\)$' -or
                $displayName -match '^PaperStream IP\s*\(TWAIN\)\s*for\s*SP\s*Series$' -or
                $displayName -match '^PaperStream IP.*TWAIN.*SP.*Series.*$' -or
                $displayName -match '^PaperStream IP.*SP.*Series.*$'
            ) {
                $category = 'TWAIN'
            }
            elseif ($displayName -match '^Software Operation Panel$') {
                $category = 'SOP'
            }
            elseif ($displayName -match '^SP Series Online Update$') {
                $category = 'OLU'
            }

            if ($null -ne $category) {
                $results.Add([pscustomobject]@{
                    Category             = $category
                    DisplayName          = $displayName
                    DisplayVersion       = if ('DisplayVersion'       -in $propNames) { [string]$item.DisplayVersion }       else { '' }
                    UninstallString      = if ('UninstallString'      -in $propNames) { [string]$item.UninstallString }      else { '' }
                    QuietUninstallString = if ('QuietUninstallString' -in $propNames) { [string]$item.QuietUninstallString } else { '' }
                    PSChildName          = if ('PSChildName'          -in $propNames) { [string]$item.PSChildName }          else { '' }
                })
            }
        }
    }

    return $results | Sort-Object Category, DisplayName, DisplayVersion, PSChildName -Unique
}

function Get-SilentUninstallCommand {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Entry,

        [Parameter(Mandatory = $true)]
        [string]$LogPath
    )

    if (-not [string]::IsNullOrWhiteSpace($Entry.QuietUninstallString)) {
        return @{
            Method  = 'QuietUninstallString'
            Command = $Entry.QuietUninstallString
        }
    }

    if ([string]::IsNullOrWhiteSpace($Entry.UninstallString)) {
        return $null
    }

    $u = $Entry.UninstallString.Trim()

    if ($u -match '(?i)\{[0-9A-Fa-f\-]{36}\}') {
        $guid = $Matches[0]
        return @{
            Method  = 'MSI ProductCode'
            Command = "msiexec.exe /x $guid /qn /norestart /L*v `"$LogPath`""
        }
    }

    if ($u -match '(?i)msiexec(\.exe)?') {
        $cmd = $u -replace '(?i)\s/I\s', ' /X '

        if ($cmd -notmatch '(?i)\s/q') {
            $cmd += " /qn /norestart /L*v `"$LogPath`""
        }
        elseif ($cmd -notmatch '(?i)/L\*v') {
            $cmd += " /L*v `"$LogPath`""
        }

        return @{
            Method  = 'MSI UninstallString'
            Command = $cmd
        }
    }

    return $null
}

function Invoke-Cmd {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CommandLine
    )

    Write-Log "Executing: $CommandLine"
    $proc = Start-Process -FilePath 'cmd.exe' -ArgumentList "/c $CommandLine" -Wait -PassThru -WindowStyle Hidden
    Write-Log "Exit code: $($proc.ExitCode)"
    return $proc.ExitCode
}

function Uninstall-Category {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Category,

        [Parameter(Mandatory = $true)]
        [string]$LogName
    )

    $entries = @(Get-UninstallEntries | Where-Object { $_.Category -eq $Category })

    if ($entries.Count -eq 0) {
        Write-Log "No installed entry found for category: $Category" 'WARN'
        return
    }

    foreach ($entry in $entries) {
        $logPath = Join-Path $LogDir $LogName
        Write-Log "Found installed entry: $($entry.DisplayName) $($entry.DisplayVersion)"

        $silent = Get-SilentUninstallCommand -Entry $entry -LogPath $logPath
        if ($null -eq $silent) {
            throw "Could not derive a silent uninstall command for '$($entry.DisplayName)'."
        }

        Write-Log "Using uninstall method: $($silent.Method)"
        $exitCode = Invoke-Cmd -CommandLine $silent.Command

        if ($exitCode -notin 0, 1641, 3010) {
            throw "Uninstall failed for '$($entry.DisplayName)' with exit code $exitCode."
        }

        if ($exitCode -in 1641, 3010) {
            Write-Log "Uninstall requested a reboot for '$($entry.DisplayName)' (exit code $exitCode)." 'WARN'
        }
    }
}

function Install-MsiPackage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$MsiPath,

        [Parameter(Mandatory = $true)]
        [string]$LogPath,

        [Parameter(Mandatory = $true)]
        [string]$FriendlyName
    )

    if (-not (Test-Path -LiteralPath $MsiPath -PathType Leaf)) {
        throw "$FriendlyName MSI not found: $MsiPath"
    }

    $args = "/i `"$MsiPath`" /qn /norestart /L*v `"$LogPath`""

    Write-Log "Installing $FriendlyName from: $MsiPath"
    Write-Log "Arguments: $args"

    $proc = Start-Process -FilePath 'msiexec.exe' -ArgumentList $args -Wait -PassThru
    Write-Log "$FriendlyName install exit code: $($proc.ExitCode)"

    if ($proc.ExitCode -notin 0, 1641, 3010) {
        throw "$FriendlyName install failed with exit code $($proc.ExitCode)."
    }

    if ($proc.ExitCode -in 1641, 3010) {
        Write-Log "$FriendlyName install requested a reboot (exit code $($proc.ExitCode))." 'WARN'
    }
}

Write-Log '========== Starting PaperStream SP reinstall =========='
Write-Log "Script directory: $ScriptDir"
Write-Log "Resolved TWAIN MSI: $TwainMsi"
Write-Log "Resolved SOP MSI: $SopMsi"
Write-Log "Resolved OLU MSI: $OluMsi"

Stop-PaperStreamProcesses

if (-not $SkipUninstall) {
    Uninstall-Category -Category 'TWAIN' -LogName 'Uninstall-TWAIN.log'

    if (-not $SkipSOP) {
        Uninstall-Category -Category 'SOP' -LogName 'Uninstall-SOP.log'
    }

    if (-not $SkipOLU) {
        Uninstall-Category -Category 'OLU' -LogName 'Uninstall-OLU.log'
    }
}
else {
    Write-Log 'Skipping uninstall by request.'
}

if (-not $SkipInstall) {
    if ([string]::IsNullOrWhiteSpace($TwainMsi)) {
        throw "Could not find the TWAIN MSI. Expected PSIP_SP_TWAIN.msi or PSIP_TWAIN.msi in $ScriptDir"
    }

    Install-MsiPackage -MsiPath $TwainMsi -LogPath $TwainInstallLog -FriendlyName 'PaperStream IP TWAIN'

    if (-not $SkipSOP) {
        if ([string]::IsNullOrWhiteSpace($SopMsi)) {
            throw "Could not find SOPSetup.msi in $ScriptDir"
        }

        Install-MsiPackage -MsiPath $SopMsi -LogPath $SopInstallLog -FriendlyName 'Software Operation Panel'
    }

    if (-not $SkipOLU) {
        if ([string]::IsNullOrWhiteSpace($OluMsi)) {
            throw "Could not find OLUSetup.msi in $ScriptDir"
        }

        Install-MsiPackage -MsiPath $OluMsi -LogPath $OluInstallLog -FriendlyName 'SP Series Online Update'
    }
}
else {
    Write-Log 'Skipping install by request.'
}

Write-Log '========== Completed =========='