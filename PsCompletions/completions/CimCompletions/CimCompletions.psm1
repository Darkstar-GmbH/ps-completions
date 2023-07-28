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
New-Module -Name CimCompletions {

    [CimCompletions] $object = $null

    function new { $this.$object = [CimCompletions]::new() }

    function Initialize { $this.$object.Initialize() }

    function GetCimInstanceCompletions { $this.$object.GetCimInstanceCompletions($args) }

    Export-ModuleMember -Function new
    Export-ModuleMember -Function Initialize
}