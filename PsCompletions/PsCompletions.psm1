# up from pwsh core 7.0 we go
#Requires -Version 7.0
#Requires -PSEdition Core

# we need this completion module as base
#Requires -Modules @{ModuleName="TabExpansionPlusPlus";ModuleVersion="1.2"}

# import all completion modules we have
using module provider/CimCompletionProvider
using module provider/WsManCompletionProvider
using module provider/FirewallCompletionProvider

Import-Module TabExpansionPlusPlus

# register all completion providers
CimCompletionProvider\Register-CimCompletionProvider
WsManCompletionProvider\Register-WsManCompletionProvider
FirewallCompletionProvider\Register-FirewallCompletionProvider
