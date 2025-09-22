# Example-RunspaceScript.ps1
# Demonstrates runspace management with module imports, custom functions, and variables

# Prerequisites: Dot source the runspace functions first
. .\Runspace-Functions.ps1

#region Setup

# Define custom function for runspaces to use (simulates system info collection)
function Get-MockSystemInfo {
    param($ComputerName, $LogPrefix)
    
    # Simulate processing time
    Start-Sleep -Seconds (Get-Random -Minimum 1 -Maximum 4)
    
    # Use Microsoft.PowerShell.Utility module cmdlets (Get-Random, Get-Date)
    $osVersions = @("Windows Server 2019", "Windows Server 2022", "Windows 11 Pro", "Windows 10 Enterprise")
    $memoryOptions = @(8, 16, 32, 64, 128)
    
    # Generate a random hash using Get-FileHash from Microsoft.PowerShell.Utility
    $tempFile = [System.IO.Path]::GetTempFileName()
    "Sample content for $ComputerName" | Out-File -FilePath $tempFile
    $fileHash = Get-FileHash -Path $tempFile -Algorithm MD5
    Remove-Item -Path $tempFile
    
    $info = @{
        ComputerName = $ComputerName
        OS = $osVersions | Get-Random
        TotalMemoryGB = $memoryOptions | Get-Random
        ProcessCount = Get-Random -Minimum 80 -Maximum 250
        SystemHash = $fileHash.Hash.Substring(0, 8)  # First 8 chars of hash
        LogMessage = "$LogPrefix - Collected system info for $ComputerName"
        Timestamp = Get-Date
        ThreadId = [System.Threading.Thread]::CurrentThread.ManagedThreadId
    }
    
    return [pscustomobject]$info
}

# Define variable for all runspaces to access
$CompanyPrefix = "ACME-CORP"

# Define list of servers to process
$ServerList = @("DC-01", "WEB-02", "DB-03", "APP-04", "FILE-05")

#endregion

#region Runspace Execution

# Create runspace pool with imported module, custom function, and variable
$runspacePool = New-RunspacePool -MaxRunspaces 3 -Modules @('Microsoft.PowerShell.Utility') -Functions @('Get-MockSystemInfo') -Variables @{CompanyPrefix = $CompanyPrefix}

# Define the script block that will run in each runspace
$scriptBlock = {
    param($ServerName)
    
    # This runs inside the runspace and has access to:
    # - Microsoft.PowerShell.Utility module (imported) - provides Get-FileHash, Get-Random, Get-Date, etc.
    # - Get-MockSystemInfo function (imported)
    # - $CompanyPrefix variable (imported)
    
    try {
        $result = Get-MockSystemInfo -ComputerName $ServerName -LogPrefix $CompanyPrefix
        return $result
    }
    catch {
        return @{
            ComputerName = $ServerName
            Error = $_.Exception.Message
            Status = "Failed"
        }
    }
}

# Create tasks for each server
$tasks = foreach ($server in $ServerList) {
    New-RunspaceTask -RunspaceId "$server" -RunspacePool $runspacePool -ScriptBlock $scriptBlock -Parameters @($server) -TaskDescription "Gathering info from $server" -TimeoutSeconds 10
}

# Wait for all tasks to complete with visual progress
Wait-RunspaceTask -Tasks $tasks -OutputType 'Visual'

# Get the results from all completed tasks
$results = Get-RunspaceResults -Tasks $tasks -IncludeMetadata

# Clean up the runspace pool
Stop-RunspacePool -RunspacePool $runspacePool

#endregion

#region Results Processing

# Display summary

# metadata first
$results | Select-Object -ExcludeProperty Results | Format-Table -AutoSize

# results
$results | Select-Object -ExpandProperty Results | Format-Table -AutoSize

#endregion
