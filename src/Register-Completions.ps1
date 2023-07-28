using assembly "C:\Windows\assembly\System.Management.Automation.dll"
using assembly "C:\Windows\assembly\System.Management.Automation.Core.dll"

using module TabExpansionPlusPlus

using module completions/CimCompletions
using module completions/WsManCompletions

$cimCompletions = [CimCompletions]::new()
$cimCompletions.Initialize()

$wsmanCompletions = [WsManCompletions]::new()
$wsmanCompletions.Initialize()

$completions = @($cimCompletions, $wsmanCompletions)
