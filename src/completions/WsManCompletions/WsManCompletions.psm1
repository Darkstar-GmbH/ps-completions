using assembly "C:\Windows\assembly\System.Management.Automation.dll"
using assembly "C:\Windows\assembly\System.Management.Automation.Core.dll"
using assembly "C:\Windows\assembly\Microsoft.Management.Infrastructure.dll"
using assembly "C:\Windows\assembly\Microsoft.WSMan.Management.dll"

using module Microsoft.WSMan.Management
using module PKI
using module CimCmdlets
using module TabExpansionPlusPlus

using namespace Microsoft.Management.Infrastructure
using namespace System.Management.Automation

<#
.SYNOPSIS
    WSMan cmdlets completions
.DESCRIPTION
    Useful completions for WSMan cmdlets
.NOTES
    nothing
.LINK
    https://github.com/.
.EXAMPLE
    Ger-WSManInstance -ResourceURI <TAB>
    Get-WSManInstance -ComputerName <TAB>
    Get-WSManInstance -CertificateThumbprint <TAB>
#>
class WsManCompletions {

    [WsManCompletions] $this = $null

    [WsManCompletions] new() {
        $this.$this = [WsManCompletions]::new()
        return $this
    }
    [void] Initialize() {
        $registerFunction = $(Get-Command -Module TabExpansionPlusPlus -Name Register-ArgumentCompleter)[0]

        & $registerFunction -CommandName "Get-WSManInstance" -ParameterName "ResourceURI" -ScriptBlock $function:GetWsManCompletions
        & $registerFunction -CommandName "Get-WSManInstance" -ParameterName "ComputerName" -ScriptBlock $function:GetWsManCompletions  
        & $registerFunction -CommandName "Get-WSManInstance" -ParameterName "CertificateThumbprint" -ScriptBlock $function:GetWsManCompletions 
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