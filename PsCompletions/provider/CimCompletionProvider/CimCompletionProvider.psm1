#Requires -Modules @{ModuleName="TabExpansionPlusPlus";ModuleVersion="1.2"}
#Requires -Modules @{ModuleName="CimCmdlets";ModuleVersion="7.0.0"}
#Requires -Modules @{ModuleName="PKI";ModuleVersion="1.0.0.0"}

using module PKI
using module CimCmdlets

using namespace System.Management.Automation

New-Module -Name CimCompletionProvider {

    <#
    .SYNOPSIS
        CIM cmdlets completions
    .DESCRIPTION
        Useful completions for CIM cmdlets
    .NOTES
        nothing
    .LINK
        https://github.com/.
    .EXAMPLE
        Get-CimInstance -ClassName <TAB>
        Get-CimInstance -CimInstance <TAB>
    #>
    function Register-CimCompletionProvider {
        
        # get module-safe register function
        $registerFunction = $(Get-Command -Module TabExpansionPlusPlus -Name Register-ArgumentCompleter)[0]

        # register provider for Get-CimInstance
        & $registerFunction -CommandName "Get-CimInstance" -ParameterName "ClassName" -ScriptBlock { CimCompletionProvider\Get-CimInstanceCompletions @args }
        & $registerFunction -CommandName "Get-CimInstance" -ParameterName "CimSession" -ScriptBlock { CimCompletionProvider\Get-CimInstanceCompletions @args }
    }

    function Get-CimInstanceCompletions {
        param ($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)

        $param = @{}

        switch -exact ($parameterName) { 

            "ClassName" {

                # evaluate bound parameters
                $ns = $fakeBoundParameter['Namespace']
                $cn = $fakeBoundParameter['ComputerName']
                $cs = $fakeBoundParameter['CimSession']

                if ($ns) { $param.Namespace = $ns }
                if ($cn) { $param.ComputerName = $cn }
                if ($cs) { $param.CimSession = $cs }

                # evaluate completion result set
                return @(CimCmdlets\Get-CimClass @param | 
                    Select-Object -ExpandProperty CimClassProperties | 
                    Select-Object -ExpandProperty Name | 
                    Where-Object { $_ -like "$wordToComplete*" } | 
                    Sort-Object | ForEach-Object {
                        New-CompletionResult `
                            -CompletionResultType ParameterValue `
                            -CompletionText $_ `
                            -ToolTip $_ `
                            -ListItemText $_
                    }
                )
            }

            "CimSession" {

                # evaluate bound parameters
                $ns = $fakeBoundParameter['Namespace']
                $cn = $fakeBoundParameter['ComputerName']
                $cl = $fakeBoundParameter['ClassName']
            
                if ($ns) { $param.Namespace = $ns }
                if ($cn) { $param.ComputerName = $cn }
                if ($cl) { $param.ClassName = $cl }
        
                # evaluate completion result set
                return @(CimCmdlets\Get-CimSession @param |
                    ForEach-Object {
                        $compound = @{}
                        $compound.Id = $_.Id
                        $compound.Session = $_
                        $compound.InstanceId = $_.InstanceId
                        $compound.Text = $_.ComputerName + ": " + $_.Id 
                        $compound
                    } | Where-Object { $_.Id -like "$wordToComplete*" } | 
                    Sort-Object -Property "Text" | ForEach-Object {
                        New-CompletionResult `
                            -CompletionResultType Variable `
                            -CompletionText $('$(' + "Get-CimSession -Id " + "$($_.Id)" + ')') `
                            -ToolTip $_.InstanceId `
                            -ListItemText $_.Text
                    }
                )
            }
        }
    }

    Export-ModuleMember -Function Get-CimInstanceCompletions
    Export-ModuleMember -Function Register-CimCompletionProvider

} | Import-Module -Global
