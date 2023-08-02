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

        # register provider for ClassName cmdlets
        Get-Command -Module "CimCmdlets" -ParameterName "ClassName" | Select-Object -ExpandProperty Name | ForEach-Object {
            & $registerFunction -CommandName $_ -ParameterName "ClassName" -ScriptBlock { CimCompletionProvider\Get-CimClassNameCompletions @args }
        }

        # register provider for CimSession cmdlets
        Get-Command -Module "CimCmdlets" -ParameterName "CimSession" | Select-Object -ExpandProperty Name | ForEach-Object {
            & $registerFunction -CommandName $_ -ParameterName "CimSession" -ScriptBlock { CimCompletionProvider\Get-CimSessionCompletions @args }
        }
    }

    function Get-CimClassNameCompletions {
        param ($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)

        $param = @{}
            
        # evaluate bound parameters
        $ns = $fakeBoundParameter['Namespace']
        $cs = $fakeBoundParameter['CimSession']

        if ($ns) { $param.Namespace = $ns }
        if ($cs) { $param.CimSession = $cs }

        # evaluate completion result set
        return @(CimCmdlets\Get-CimClass @param -ClassName "*$wordToComplete*" | 
            Select-Object -ExpandProperty CimClassName | 
            Where-Object { $_ -like "*$wordToComplete*" } | 
            Sort-Object | ForEach-Object {
                New-CompletionResult `
                    -CompletionResultType ParameterValue `
                    -CompletionText $_ `
                    -ToolTip $_ `
                    -ListItemText $_
            }
        )
    }

    function Get-CimSessionCompletions {
        param ($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)

        # evaluate bound parameters
        $ns = $fakeBoundParameter['Namespace']
    
        if ($ns) { $param.Namespace = $ns }

        # evaluate completion result set
        return @(CimCmdlets\Get-CimSession @param |
            ForEach-Object {
                $compound = @{}
                $compound.Id = $_.Id
                $compound.Session = $_
                $compound.InstanceId = $_.InstanceId
                $compound.ComputerName = $_.ComputerName
                $compound.Text = $_.ComputerName + ": " + $_.Id 
                $compound
            } | Where-Object { $_.Text -like "*$wordToComplete*" } | 
            Sort-Object -Property "Text" | ForEach-Object {
                New-CompletionResult `
                    -CompletionResultType Variable `
                    -CompletionText $('$(' + "Get-CimSession -Id " + "$($_.Id)" + ')') `
                    -ToolTip $_.InstanceId `
                    -ListItemText $_.Text
            }
        )
    }

    Export-ModuleMember -Function Register-CimCompletionProvider

    Export-ModuleMember -Function Get-CimClassNameCompletions
    Export-ModuleMember -Function Get-CimSessionCompletions

} | Import-Module -Global
