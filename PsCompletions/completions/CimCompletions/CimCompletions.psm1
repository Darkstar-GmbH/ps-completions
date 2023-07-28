#Requires -Modules @{ModuleName="TabExpansionPlusPlus";ModuleVersion="1.2"}
#Requires -Modules @{ModuleName="CimCmdlets";ModuleVersion="7.0.0"}
#Requires -Modules @{ModuleName="PKI";ModuleVersion="1.0.0.0"}

using namespace Microsoft.Management.Infrastructure
using namespace System.Management.Automation

Import-Module PKI
Import-Module CimCmdlets
Import-Module TabExpansionPlusPlus

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

class CimCompletions {

    [CimCompletions] $this = $null

    [CimCompletions] new() {
        $this.$this = [CimCompletions]::new()
        return $this
    }

    [void] Initialize() {
        $registerFunction = $(Get-Command -Module TabExpansionPlusPlus -Name Register-ArgumentCompleter)[0]

        & $registerFunction -CommandName "Get-CimInstance" -ParameterName "ClassName" -ScriptBlock { 
            param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
            return $this.GetCimInstanceCompletions($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
        }

        & $registerFunction -CommandName "Get-CimInstance" -ParameterName "CimInstance" -ScriptBlock {
            param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
            return $this.GetCimInstanceCompletions($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
        }
    }

    [CompletionResult[]] GetCimInstanceCompletions (
        [string] $commandName, 
        [string] $parameterName, 
        [string] $wordToComplete,
        [string] $commandAst, 
        [string] $fakeBoundParameters) {
            
        switch (-exact "$parameterName") { 
            "ClassName" {
                return @(Get-CimClass | Where-Object { ($wordToComplete + "*") -like $_.CimSystemProperties.ClassName } | ForEach-Object {
                        New-CompletionResult `
                            -CompletionText $_.CimSystemProperties.ClassName `
                            -ToolTip $_.CimSystemProperties.ClassName `
                            -Description $_.CimSystemProperties.ClassName
                    })
            }

            "CimInstance" {
                return @(Get-CimSession | Where-Object { ($wordToComplete + "*") -like $_.InstanceId } | ForEach-Object {
                        New-CompletionResult `
                            -CompletionText $_ `
                            -ToolTip $_.InstanceId `
                            -Description $_.InstanceId 
                    })
            }

            default { break; }
        }

        return @();
    }
}