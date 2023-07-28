#Requires -Modules @{ModuleName="Microsoft.WSMan.Management";ModuleVersion="7.0.0"}
#Requires -Modules @{ModuleName="PKI";ModuleVersion="1.0.0.0"}
#Requires -Modules @{ModuleName="CimCmdlets";ModuleVersion="7.0.0"}
#Requires -Modules @{ModuleName="TabExpansionPlusPlus";ModuleVersion="1.2"}

using module Microsoft.WSMan.Management
using module PKI
using module CimCmdlets

using namespace Microsoft.Management.Infrastructure
using namespace System.Management.Automation

Import-Module TabExpansionPlusPlus

class WsManCompletions {

    hidden [WsManCompletions] $this = $null

    [WsManCompletions] new() {
        $this.$this = [WsManCompletions]::new()
        return $this
    }
    
    [void] Initialize() {
        $registerFunction = $(Get-Command -Module TabExpansionPlusPlus -Name Register-ArgumentCompleter)[0]

        & $registerFunction -CommandName "Get-WSManInstance" -ParameterName "ResourceURI" -ScriptBlock {
            param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
            return $this.GetWsManCompletions($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
        }

        & $registerFunction -CommandName "Get-WSManInstance" -ParameterName "ComputerName" -ScriptBlock { 
            param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
            return $this.GetWsManCompletions($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
        }

        & $registerFunction -CommandName "Get-WSManInstance" -ParameterName "CertificateThumbprint" -ScriptBlock {
            param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
            return $this.GetWsManCompletions($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
        } 
    }

    [CompletionResult[]] GetWsManCompletions(
        [string] $commandName, 
        [string] $parameterName, 
        [string] $wordToComplete,
        [string] $commandAst, 
        [string] $fakeBoundParameters
    ) {

        switch (-exakt "$parameterName") { 
            "ResourceURI" {
                return @(Get-CimClass | Where-Object { ("$wordToComplete" + "*") -like $_.CimSystemProperties.ClassName } | ForEach-Object {
                        New-CompletionResult `
                            -CompletionText "wmi/root/cimv2/" + $_.CimSystemProperties.ClassName `
                            -ToolTip $_.CimSystemProperties.ClassName
                    }) 
            }

            "ComputerName" { 
                return @(Get-CimInstance -ClassName CIM_ComputerSystem | Where-Object { ("$wordToComplete" + "*") -like $_.Name } | ForEach-Object { 
                        New-CompletionResult `
                            -CompletionText $_.Name `
                            -ToolTip $_.Name + "(" + $_.PrimaryOwnerName + ")"
                    })
            }
        
            "CertificateThumbprint" {
                return @(Get-ChildItem -Path Cert:\ -Recurse | Where-Object { ("$wordToComplete" + "*") -like $_.Thumbprint } | ForEach-Object { 
                        New-CompletionResult `
                            -CompletionText $_.Thumbprint `
                            -ToolTip $_.IssuerName + "|" $_.Subject + "|" + $_.DnsNameList + "(" + $_.Thumbprint + ")"
                    })
            }

            default { break }
        }

        return @()
    }
}
