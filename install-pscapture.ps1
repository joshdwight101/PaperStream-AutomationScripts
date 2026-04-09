#Requires -RunAsAdministrator
<#
Author : Joshua Dwight
Github : https://github.com/joshdwight101

Purpose:
- Silent installer for PaperStream Capture only
- Expects this script to be in the main folder
- Expects PSCapture files in the subfolder:
    .\PSCapture\PSCSetup.exe
    .\PSCapture\Data1\setup_en.msi
#>

[CmdletBinding()]
param(
    [string]$InstallDir = 'C:\Program Files\fiScanner\PaperStream Capture'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ScriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$PscRoot     = Join-Path $ScriptDir 'PSCapture'
$PscSetupExe = Join-Path $PscRoot 'PSCSetup.exe'
$SetupMsi    = Join-Path $PscRoot 'Data1\setup_en.msi'
$LogDir      = Join-Path $ScriptDir 'Logs'
$MainLog     = Join-Path $LogDir 'Install-PSCapture.log'

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

function Stop-PSCaptureProcesses {
    $processNames = @(
        'PaperStreamCapture',
        'PSCSetup',
        'TWUNK_32',
        'TWUNK_64'
    )

    foreach ($name in $processNames) {
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

function Test-InstallFiles {
    if (-not (Test-Path -LiteralPath $PscRoot -PathType Container)) {
        throw "PSCapture folder was not found in: $ScriptDir"
    }

    if (-not (Test-Path -LiteralPath $PscSetupExe -PathType Leaf)) {
        throw "PSCSetup.exe was not found in: $PscRoot"
    }

    if (-not (Test-Path -LiteralPath $SetupMsi -PathType Leaf)) {
        throw "setup_en.msi was not found in: $(Split-Path -Parent $SetupMsi)"
    }
}

function Install-PSCapture {
    $arguments = "`"$SetupMsi`" INSTALLDIR=`"$InstallDir`" -q -b"

    Write-Log "Using PSCapture root: $PscRoot"
    Write-Log "Using PSCSetup.exe: $PscSetupExe"
    Write-Log "Using MSI: $SetupMsi"
    Write-Log "Install directory: $InstallDir"
    Write-Log "Arguments: $arguments"

    $proc = Start-Process -FilePath $PscSetupExe -ArgumentList $arguments -Wait -PassThru
    Write-Log "Exit code: $($proc.ExitCode)"

    if ($proc.ExitCode -notin 0, 1641, 3010) {
        throw "PaperStream Capture install failed with exit code $($proc.ExitCode)."
    }

    if ($proc.ExitCode -eq 0) {
        Write-Log 'PaperStream Capture installed successfully.'
    }
    elseif ($proc.ExitCode -in 1641, 3010) {
        Write-Log "Installation completed and a reboot is required. Exit code: $($proc.ExitCode)" 'WARN'
    }
}

Write-Log '========== Starting PaperStream Capture install =========='
Write-Log "Script directory: $ScriptDir"

Test-InstallFiles
Stop-PSCaptureProcesses
Install-PSCapture

Write-Log '========== Completed =========='