# Registeres all completion handler with the powershell engine, that is, with the module TabExpansionPlusPlus

# up from pwsh core 7.0 we go
#Requires -Version 7.0
#Requires -PSEdition Core

# we need this completion module as base
#Requires -Modules @{ModuleName="TabExpansionPlusPlus";ModuleVersion="1.2"}

# import all completion modules we have
using module completions/CimCompletions
using module completions/WsManCompletions

using namespace cimcomp::CimCompletions
using namespace wsmancomp::WsManCompletions

# this is the main class that registers all completion handlers
class PsCompletions {

    hidden [PsCompletions] new() {
        $this.$this = [PsCompletions]::new()
        return $this
    }

    hidden [void] RegisterCompletions() {
        $cimCompletions = [CimCompletions]::new()
        $cimCompletions.Initialize()
        
        $wsmanCompletions = [WsManCompletions]::new()
        $wsmanCompletions.Initialize()
    }
}

function RegisterCompletions {
    $psCompletions = [PsCompletions]::new()
    $psCompletions.RegisterCompletions()
}

# eventually, register all completion handler
RegisterCompletions

