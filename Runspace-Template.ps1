# ===== PARALLEL TASK TEMPLATE =====

# 1. import runspace function library
. .\Runspace-Functions.ps1

# 2. Set up your items
$items = @("item1", "item2", "item3", "item4", "item5")

function Test-RunspaceFunction {
    param($InputObject)

    "$Header $InputObject received"
}

# 3. Define your script block (use param() for the item)
$scriptBlock = {
    param($Item)

    # Your code here

    # simulating work
    $randDuration = Get-Random -Minimum 2 -Maximum 5
    Start-Sleep -Seconds $randDuration

    # output results
    Test-RunspaceFunction $Item
}

# 4. Choose what to import
$modules = @('Microsoft.PowerShell.Utility') # Utility might cause a false warning
$functions = @('Test-RunspaceFunction')
$variables = @{Header = '[INFO]'}

# 5. Run it
$pool = New-RunspacePool -MaxRunspaces 5 -Modules $modules -Functions $functions -Variables $variables
$tasks = foreach ($item in $items) {
    New-RunspaceTask -RunspacePool $pool -ScriptBlock $scriptBlock -Parameters @($item) -TimeoutSeconds 60
}
Wait-RunspaceTask -Tasks $tasks -OutputType Visual
$results = Get-RunspaceResults -Tasks $tasks
Stop-RunspacePool -RunspacePool $pool

# 6. output results
$results
# ===== END TEMPLATE =====