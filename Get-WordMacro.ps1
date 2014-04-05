function Get-WordMacro
{
<#
.SYNOPSIS

Outputs the contents on a Word macro if it is present in an Word document.

Author: Matthew Graeber (@mattifestation)
License: BSD 3-Clause

.DESCRIPTION

Get-WordMacro outputs the contents of an Word macro if it is present and
optionally removes it if the '-Remove' switch is provided.

.PARAMETER Remove

Remove all macros from a Word document if they are present.

.PARAMETER Force

Suppress the confirmation prompt when the '-Remove' switch is used.

.EXAMPLE

Get-ChildItem C:\* -Recurse -Include "*.doc","*.docx" | Get-WordMacro

.EXAMPLE

Get-WordMacro -Path evil.doc -Remove

.NOTES

Get-WordMacro relies on the Word COM object which requires that Word be
installed. When the '-Remove' switch is provided to Get-WordMacro, all macros
are removed from the Word document. Be mindful of this fact if your
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
            $Word = New-Object -ComObject Word.Application
        }
        catch
        {
            throw 'Word is not installed. Get-WordMacro requires that Word be installed.'
        }

        # EXTREMELY IMPORTANT!!!
        # Disable automatic execution of macros
        $Word.AutomationSecurity = 'msoAutomationSecurityForceDisable'
        
        # Disable save dialogs
        $Word.DisplayAlerts = 'wdAlertsNone'

        # Macro security settings need to be briefly disabled in order to view and remove macros.
        # They will be restored at the end of the script
        Set-ItemProperty HKCU:\Software\Microsoft\Office\*\*\Security -Name AccessVBOM -Type DWORD -Value 1
        Set-ItemProperty HKCU:\Software\Microsoft\Office\*\*\Security -Name VBAWarnings -Type DWORD -Value 1
    }

    PROCESS
    {
        foreach ($File in $Path)
        {
            $FullPath = Resolve-Path $File
            
            $Document = $Word.Documents.Open($FullPath.Path)
            $CodeModule = $Document.VBProject.VBComponents.Item(1).CodeModule

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
                            WordDocument = $FullPath
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

                    # http://msdn.microsoft.com/en-us/library/ff839952(v=office.14).aspx
                    $wdFormatDocumentDefault = 16
                    $Document.SaveAs([Ref] $FullPath.Path, [Ref] $wdFormatDocumentDefault)

                    if ($MacroObject)
                    {
                        Write-Verbose "Removed macro from $FullPath."
                    }
                }
            }

            $Document.Close()
        }
    }

    END
    {
        # Re-enable macro security settings
        Set-ItemProperty HKCU:\Software\Microsoft\Office\*\*\Security -Name AccessVBOM -Type DWORD -Value 0
        Set-ItemProperty HKCU:\Software\Microsoft\Office\*\*\Security -Name VBAWarnings -Type DWORD -Value 0

        $Word.Quit()

        # Kill orphaned Word process
        Get-Process winword | ? { $_.MainWindowHandle -eq 0 } | Stop-Process
    }
}
