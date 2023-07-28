#Requires -Modules @{ModuleName="Microsoft.WSMan.Management";ModuleVersion="7.0.0"}
#Requires -Modules @{ModuleName="PKI";ModuleVersion="1.0.0.0"}
#Requires -Modules @{ModuleName="CimCmdlets";ModuleVersion="7.0.0"}
#Requires -Modules @{ModuleName="TabExpansionPlusPlus";ModuleVersion="1.2"}

using module Microsoft.WSMan.Management
using module PKI
using module CimCmdlets

New-Module -Name WsManCompletionProvider {

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
    function Register-WsManCompletionProvider {
        
        # get module-safe register function
        $registerFunction = $(Get-Command -Module TabExpansionPlusPlus -Name Register-ArgumentCompleter)[0]

        # register provider for Get-CimInstance
        & $registerFunction -CommandName "Get-WSManInstance" -ParameterName "ResourceURI" -ScriptBlock { WsManCompletionProvider\Get-WsManCompletions $args }
        & $registerFunction -CommandName "Get-WSManInstance" -ParameterName "ComputerName" -ScriptBlock { WsManCompletionProvider\Get-WsManCompletions $args }
        & $registerFunction -CommandName "Get-WSManInstance" -ParameterName "CertificateThumbprint" -ScriptBlock { WsManCompletionProvider\Get-WsManCompletions $args }
    }

    function Get-WsManCompletions {
        param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)

        switch -exact ($parameterName) { 
             
            "ResourceURI" {
                Get-CimClass | Where-Object { ("$wordToComplete" + "*") -like $_.CimSystemProperties.ClassName } | ForEach-Object {
                    New-CompletionResult `
                        -CompletionText "wmi/root/cimv2/" + $_.CimSystemProperties.ClassName `
                        -ToolTip $_.CimSystemProperties.ClassName
                }
            }

            "ComputerName" { 
                Get-CimInstance -ClassName CIM_ComputerSystem | Where-Object { ("$wordToComplete" + "*") -like $_.Name } | ForEach-Object { 
                    New-CompletionResult `
                        -CompletionText $_.Name `
                        -ToolTip $_.Name + "(" + $_.PrimaryOwnerName + ")"
                }
            }
        
            "CertificateThumbprint" {
                Get-ChildItem -Path Cert:\ -Recurse | Where-Object { ("$wordToComplete" + "*") -like $_.Thumbprint } | ForEach-Object { 
                    New-CompletionResult `
                        -CompletionText $_.Thumbprint `
                        -ToolTip $_.IssuerName + "|" $_.Subject + "|" + $_.DnsNameList + "(" + $_.Thumbprint + ")"
                }
            }
        }
    }

    Export-ModuleMember -Function Get-WsManCompletions
    Export-ModuleMember -Function Register-WsManCompletionProvider

} | Import-Module -Global
