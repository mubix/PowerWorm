function Test-PowerWormInfection
{
<#
.SYNOPSIS

Detects the presence of a Power Worm infection.

Author: Matthew Graeber (@mattifestation)
License: BSD 3-Clause

.DESCRIPTION

Test-PowerWormInfection will alert you if it detects any file or registry
artifacts left behind by Power Worm. You can also optionally remove
Power Worm artifacts from an infected system.

.PARAMETER Remove

Deletes Power Worm file and registry artifacts if they are present.

.EXAMPLE

Test-PowerWormInfection

.EXAMPLE

Test-PowerWormInfection -Remove

.NOTES

Test-PowerWormInfection does not remove malicious macros from Office
documents. To remove macros from Office documents, use the Get-ExcelMacro
and Get-WordMacro functions with the '-Remove' switch.

.LINK

http://www.exploit-monday.com/2014/04/powerworm-analysis.html
#>

    [CmdletBinding(SupportsShouldProcess=$True, ConfirmImpact='High')] Param ([Switch] $Remove)

    $MachineGUID = (Get-WmiObject Win32_ComputerSystemProduct).UUID

    $DownloadedToolsPath = Join-Path $Env:APPDATA $MachineGuid
    $DownloadedToolsDir = $null

    $Infected = $False

    Write-Verbose 'Testing for Power Worm file artifacts...'

    if (Test-Path $DownloadedToolsPath)
    {
        Write-Warning 'Power Worm may have executed based upon the existence of the following path:'
        Write-Warning $DownloadedToolsPath

        $DownloadedToolsDir = Get-ChildItem -LiteralPath (Split-Path -Parent $DownloadedToolsPath) -Attributes D+H+S+NotContentIndexed
    }

    $Files = Get-ChildItem -Force $DownloadedToolsPath -ErrorAction SilentlyContinue

    if ($Files)
    {
        Write-Warning 'The following files should be deleted:'

        foreach ($File in $Files)
        {
            Write-Warning $File.FullName
        }
    }

    Write-Verbose 'Testing for Power Worm registry artifacts...'

    $Payload1Properties = @{
        Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
        Name = $MachineGuid
        ErrorAction = 'SilentlyContinue'
    }

    $Payload2Properties = @{
        Path = 'HKCU:\Software\Microsoft'
        Name = ($MachineGuid + '0')
        ErrorAction = 'SilentlyContinue'
    }

    $Payload3Properties = @{
        Path = 'HKCU:\Software\Microsoft'
        Name = ($MachineGuid + '1')
        ErrorAction = 'SilentlyContinue'
    }

    $Payload1 = Get-ItemProperty @Payload1Properties
    $Payload2 = Get-ItemProperty @Payload2Properties
    $Payload3 = Get-ItemProperty @Payload3Properties

    if ($Payload1)
    {
        Write-Warning "A Power Worm payload was found in $($Payload1Properties['Path']) -> $($Payload1Properties.Name)"
    }

    if ($Payload2)
    {
        Write-Warning "A Power Worm payload was found in $($Payload2Properties['Path']) -> $($Payload2Properties.Name)"
    }

    if ($Payload3)
    {
        Write-Warning "A Power Worm payload was found in $($Payload3Properties['Path']) -> $($Payload3Properties.Name)"
    }

    if (-not $Remove)
    {
        # If there are any Power Worm artifacts, then you are most likely infected
        if ($DownloadedToolsDir -or $Files -or $Payload1 -or $Payload2 -or $Payload3)
        {
            New-Object PSObject -Property @{ Infected = $True }
        }
        else
        {
            New-Object PSObject -Property @{ Infected = $False }
        }
    }
    else
    {
        if ($Files -or $DownloadedToolsDir)
        {
            Write-Verbose 'Removing discovered Power Worm file artifacts...'

            if ($Files)
            {
                $Files | Remove-Item
            }

            if ($DownloadedToolsDir)
            {
                # Restore normal directory attributes
                $DownloadedToolsDir.Attributes = 'Directory'
                $DownloadedToolsDir | Remove-Item
            }
        }

        if ($Payload1 -or $Payload2 -or $Payload3)
        {
            Write-Verbose 'Removing discovered Power Worm registry artifacts...'

            if ($Payload1) { Remove-ItemProperty @Payload1Properties }
            if ($Payload2) { Remove-ItemProperty @Payload2Properties }
            if ($Payload3) { Remove-ItemProperty @Payload3Properties }
        }

        Write-Verbose 'Power Worm artifacts removed. Office documents may remain infected though.'
    }
}
