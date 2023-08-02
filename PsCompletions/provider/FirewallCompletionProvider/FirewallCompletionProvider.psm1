#Requires -Modules @{ModuleName="TabExpansionPlusPlus";ModuleVersion="1.2"}
#Requires -Modules @{ModuleName="NetSecurity";ModuleVersion="2.0.0.0"}
#Requires -Modules @{ModuleName="CimCmdlets";ModuleVersion="7.0.0"}
#Requires -Modules @{ModuleName="CimCompletionProvider";ModuleVersion="0.1.0"}

using module NetSecurity
using module CimCmdlets

<#
.SYNOPSIS
 netsecurity module firewall completion provider
.DESCRIPTION
 This completion provider comes with full support for all NetSecurity Firewall cmdlets.
.NOTES
 pre-release version
.LINK
 https://github.com/Darkstar-GmbH/ps-completions
.COMPONENT
 ps-completions
.FUNCTIONALITY
 provider/firewall
.EXAMPLE
 Get-NetFirewallRule -AssociatedNetFirewallAddressFilter <TAB>
#>
New-Module -Name FirewallCompletionProvider {

    <#
    .SYNOPSIS
     CIM configuration provider
    .DESCRIPTION
     Provides configuration for matching Cmdlets/ParameterNames with CIM-Instances.
    .NOTES
     pre-release version
    .LINK
     https://github.com/Darkstar-GmbH/ps-completions
    .COMPONENT
     ps-completions
    .FUNCTIONALITY
     provider/firewall/configuration
    .EXAMPLE
     GetCimConfiguration
    #>
    function Get-CimConfiguration {
        return @(
            [pscustomobject] @{
                ParameterName = [string] "AssociatedNetFirewallAddressFilter" 
                CimClass      = [string] "MSFT_NetAddressFilter" 
            },    
            [pscustomobject] @{
                ParameterName = [string] "AssociatedNetFirewallApplicationFilter"
                CimClass      = [string] "MSFT_NetApplicationFilter" 
            },
            [pscustomobject] @{
                ParameterName = [string] "AssociatedNetFirewallInterfaceFilter"
                CimClass      = [string] "MSFT_NetInterfaceFilter" 
            },
            [pscustomobject] @{
                ParameterName = [string] "AssociatedNetFirewallInterfaceTypeFilter" 
                CimClass      = [string] "MSFT_NetInterfaceTypeFilter" 
            },
            [pscustomobject] @{
                ParameterName = [string] "AssociatedNetFirewallPortFilter"
                CimClass      = [string] "MSFT_NetPortFilter" 
            },
            [pscustomobject] @{
                ParameterName = [string] "AssociatedNetFirewallSecurityFilter"
                CimClass      = [string] "MSFT_NetSecurityFilter" 
            },
            [pscustomobject] @{
                ParameterName = [string] "AssociatedNetFirewallServiceFilter"
                CimClass      = [string] "MSFT_NetServiceFilter" 
            },
            [pscustomobject] @{
                ParameterName = [string] "AssociatedNetFirewallProfile"
                CimClass      = [string] "MSFT_NetProfile" 
            }
        )
    }

    <#
    .SYNOPSIS
     Registeres all completions for the firewall cmdlets.
    .DESCRIPTION
     Registeres all completion handler that are neccessary to provide the ps-completions usability.
    .NOTES
     pre-release version
    .LINK
     https://github.com/Darkstar-GmbH/ps-completions
    .COMPONENT
     ps-completions
    .FUNCTIONALITY
     provider/firewall/register
    .EXAMPLE
     Register-FirewallCompletionProvider
    #>
    function Register-FirewallCompletionProvider {
        
        # get module-safe register function
        $registerFunction = $(Get-Command -Module TabExpansionPlusPlus -Name Register-ArgumentCompleter)[0]

        # register provider for AssociatedNetFirewallRule
        Get-Command -Module "NetSecurity" -ParameterName "AssociatedNetFirewallRule" | Select-Object -ExpandProperty Name | ForEach-Object {
            & $registerFunction -CommandName $_ -ParameterName "AssociatedNetFirewallRule" -ScriptBlock { FirewallCompletionProvider\Get-FirewallAssociatedRuleCompletions @args }
        }

        # register cim class completions
        Foreach ($object in Get-CimConfiguration) {
            Get-Command -Module "NetSecurity" -ParameterName $object.ParameterName | ForEach-Object {
                & $registerFunction -CommandName $_.Name -ParameterName $object.ParameterName -ScriptBlock { FirewallCompletionProvider\Get-FirewallAssociatedCIMCompletions @args }
            }
        }
 
        # register CIM session handling
        Get-Command -Module "NetSecurity" -ParameterName "CimSession" | Select-Object -ExpandProperty Name | ForEach-Object {
            & $registerFunction -CommandName $_ -ParameterName "CimSession" -ScriptBlock { CimCompletionProvider\Get-CimSessionCompletions @args }
        }
    }

    <#
    .SYNOPSIS
    The Firewall AssociatedFirewallRule completion provider.
    .DESCRIPTION
    Provides completions for all cmdlets that have the parameter AssociatedFirewallRule.
    .NOTES
    pre-release version
    .LINK
    https://github.com/Darkstar-GmbH/ps-completions
    .COMPONENT
    ps-completions
    .FUNCTIONALITY
    provider/firewall/associatedrule
    .EXAMPLE
    Get-NetFirewallAddressFilter -AssociatedNetFirewallRule <TAB>
    #>
    function Get-FirewallAssociatedRuleCompletions {

        [CmdletBinding()]
        param ($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)

        # evaluate completion result set
        return @(NetSecurity\Get-NetFirewallRule -Name "*$wordToComplete*" | 
            Select-Object -ExpandProperty Name | 
            Where-Object { $_ -like "*$wordToComplete*" } | 
            Sort-Object | ForEach-Object {
                New-CompletionResult `
                    -CompletionResultType Variable `
                    -CompletionText $('$(' + "Get-NetFirewallRule -Name " + "$($_)" + ')') `
                    -ToolTip $_ `
                    -ListItemText $_
            }
        )
    }

    <#
    .SYNOPSIS
    The firewall AssociatedFirewall* completion provider
    .DESCRIPTION
    Provides completions for all cmdlets that have parameters AssociatedFirewall*
    .NOTES
    pre-release version
    .LINK
    https://github.com/Darkstar-GmbH/ps-completions
    .COMPONENT
    ps-completions
    .FUNCTIONALITY
    provider/firewall/associatedcim
    .EXAMPLE
    Get-NetFirewallRule -AssociatedNetFirewallApplicationFilter <TAB>
    #>
    function Get-FirewallAssociatedCIMCompletions {
        
        [CmdletBinding()]
        param ($CommandName, $ParameterName, $WordToComplete, $CommandAst, $FakeBoundParameter)

        $CimClass = FirewallCompletionProvider\Get-CimConfiguration | 
        Where-Object { $_.ParameterName -eq "${ParameterName}" } | 
        Select-Object -ExpandProperty "CimClass" |
        Select-Object -First 1

        # evaluate completion result set
        return @(CimCmdlets\Get-CimInstance -Namespace "root/standardcimv2" -ClassName "${CimClass}" | 
            ForEach-Object {
                $compound = @{}
                $compound.Instance = $_
                $compound.InstanceId = $($_ | 
                    Select-Object -ExpandProperty CimInstanceProperties | 
                    Where-Object -Property Name -eq "InstanceID" | 
                    Select-Object -ExpandProperty Value)
                $compound
            } | Where-Object { $_.InstanceId -like "*${WordToComplete}*" } | 
            Sort-Object | ForEach-Object {
                New-CompletionResult `
                    -CompletionResultType Variable `
                    -CompletionText $('$(' + "Get-CimInstance -Namespace ""root/standardcimv2"" -Query ""SELECT * FROM ${CimClass} WHERE InstanceID like '%" + $($_.InstanceId) + "%'""" + ')') `
                    -ToolTip $_.InstanceId `
                    -ListItemText $_.InstanceId
            }
        )
    }

    Export-ModuleMember -Function Register-FirewallCompletionProvider

    Export-ModuleMember -Function Get-FirewallAssociatedRuleCompletions
    Export-ModuleMember -Function Get-FirewallAssociatedCIMCompletions

    Export-ModuleMember -Function Get-CimConfiguration

} | Import-Module -Global
