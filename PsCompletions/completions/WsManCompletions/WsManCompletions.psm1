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
New-Module -Name WsManCompletions {

    [WsManCompletions] $object = $null

    function new { $this.$object = [WsManCompletions]::new() }
    
    function Initialize { $this.$object.Initialize() }

    function GetWsManCompletions { $this.$object.GetWsManCompletions($args) }

    Export-ModuleMember -Function new
    Export-ModuleMember -Function Initialize
}
