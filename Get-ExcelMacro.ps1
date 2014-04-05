function Get-ExcelMacro
{
<#
.SYNOPSIS

Outputs the contents on an Excel macro if it is present in an Excel spreadsheet.

Author: Matthew Graeber (@mattifestation)
License: BSD 3-Clause

.DESCRIPTION

Get-ExcelMacro outputs the contents of an Excel macro if it is present and
optionally removes it if the '-Remove' switch is provided.

.PARAMETER Remove

Remove all macros from an Excel spreadsheet if they are present.

.PARAMETER Force

Suppress the confirmation prompt when the '-Remove' switch is used.

.EXAMPLE

Get-ChildItem C:\* -Recurse -Include "*.xls","*.xlsx" | Get-ExcelMacro

.EXAMPLE

Get-ExcelMacro -Path evil.xls -Remove

.NOTES

Get-ExcelMacro relies on the Excel COM object which requires that Excel be
installed. When the '-Remove' switch is provided to Get-ExcelMacro, all macros
are removed from the Excel spreadsheet. Be mindful of this fact if your
organization relies upon macros for legitimate purposes.

.LINK

http://www.exploit-monday.com/2014/04/powerworm-analysis.html
#>

    [CmdletBinding(SupportsShouldProcess = $True , ConfirmImpact = 'Medium')]
    Param (
        [Parameter(Position = 0, ValueFromPipeline = $True)]
        [ValidateScript({Test-Path $_})]
        [String[]]
        $Path,

        [Parameter(ParameterSetName = 'Remove')]
        [Switch]
        $Remove,

        [Parameter(ParameterSetName = 'Remove')]
        [Switch]
        $Force
    )
    
    BEGIN
    {
        try
        {
            $Excel = New-Object -ComObject Excel.Application
        }
        catch
        {
            throw 'Excel is not installed. Get-ExcelMacro requires that Excel be installed.'
        }

        # EXTREMELY IMPORTANT!!!
        # Disable automatic execution of macros
        $Excel.AutomationSecurity = 'msoAutomationSecurityForceDisable'
        
        # Disable save dialogs
        $Excel.DisplayAlerts = $False

        # Macro security settings need to be briefly disabled in order to view and remove macros.
        # They will be restored at the end of the script
        Set-ItemProperty HKCU:\Software\Microsoft\Office\*\*\Security -Name AccessVBOM -Type DWORD -Value 1
        Set-ItemProperty HKCU:\Software\Microsoft\Office\*\*\Security -Name VBAWarnings -Type DWORD -Value 1
    }

    PROCESS
    {
        foreach ($File In $Path)
        {
            $FullPath = Resolve-Path $File
    
            $Workbooks = $Excel.Workbooks.Open($FullPath.Path)
            $CodeModule = $Workbooks.VBProject.VBComponents.Item(1).CodeModule

            $CurrentLine = 1

            if ($CodeModule)
            {
                do
                {
                    $FunctionName = $CodeModule.ProcOfLine($CurrentLine, [ref] 0)

                    if ($FunctionName -ne $null)
                    {
                        $StartLine = $CurrentLine
                        $EndingLine = $CodeModule.ProcCountLines($FunctionName, 0) - 1

                        $Result = @{
                            ExcelDocument = $FullPath
                            FunctionName = $FunctionName
                            MacroContents = $CodeModule.Lines($StartLine, $EndingLine)
                        }

                        $MacroObject = New-Object PSObject -Property $Result
                        $MacroObject

                        $CurrentLine += $EndingLine
                    }
                } while ($CurrentLine -lt $CodeModule.CountOfLines)

                if ($Remove)
                {
                    if ($Force -or ($Response = $PSCmdlet.ShouldContinue('Do you want to proceed?',
                        "Removing macro from $FullPath"))) { }

                    if ($CodeModule.CountOfLines -gt 0)
                    {
                        $CodeModule.DeleteLines(1, $CodeModule.CountOfLines)
                    }

                    # http://msdn.microsoft.com/en-us/library/ff198017(v=office.14).aspx
                    $xlWorkbookDefault = 51

                    $Workbooks.SaveAs($FullPath.Path, $xlWorkbookDefault)

                    if ($MacroObject)
                    {
                        Write-Verbose "Removed macro from $FullPath."
                    }
                }
            }

            $Workbooks.Close($False)
        }
    }

    END
    {
        # Re-enable macro security settings
        Set-ItemProperty HKCU:\Software\Microsoft\Office\*\*\Security -Name AccessVBOM -Type DWORD -Value 0
        Set-ItemProperty HKCU:\Software\Microsoft\Office\*\*\Security -Name VBAWarnings -Type DWORD -Value 0

        $Excel.Quit()

        # Kill orphaned Excel process
        Get-Process excel | ? { $_.MainWindowHandle -eq 0 } | Stop-Process
    }
}

