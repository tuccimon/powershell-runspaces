# Runspace Management Module

#region Core Runspace Functions

function New-RunspaceSessionState {
    <#
    .SYNOPSIS
    Creates an InitialSessionState for runspaces with custom functions and modules.
    
    .PARAMETER Functions
    Array of function names to import from current session
    
    .PARAMETER Modules  
    Array of module names to import
    
    .PARAMETER Variables
    Hashtable of variables to add to session state
    
    .PARAMETER RestrictedCommands
    Use minimal command set instead of default PowerShell commands
    #>
    [CmdletBinding()]
    param(
        [System.Collections.Generic.List[string]]$Functions = [System.Collections.Generic.List[string]]::new(),
        [System.Collections.Generic.List[string]]$Modules = [System.Collections.Generic.List[string]]::new(),
        [hashtable]$Variables = @{},
        [switch]$RestrictedCommands
    )
    
    # Create initial session state - default is full commands unless restricted
    if ($RestrictedCommands) {
        $initialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault2()
    }
    else {
        $initialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    }
    
    # Add custom functions
    foreach ($FunctionName in $Functions) {
        try {
            $functionItem = Get-Item "function:$FunctionName" -ErrorAction Stop
            $functionEntry = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry(
                $FunctionName, 
                $functionItem.Definition
            )
            [void]$initialSessionState.Commands.Add($functionEntry)
            Write-Verbose "Added function: $FunctionName"
        }
        catch {
            Write-Warning "Failed to add function '$FunctionName': $($_.Exception.Message)"
        }
    }
    
    # Add modules
    foreach ($ModuleName in $Modules) {
        try {
            [void]$initialSessionState.ImportPSModule($ModuleName)
            Write-Verbose "Added module: $ModuleName"
        }
        catch {
            Write-Warning "Failed to add module '$ModuleName': $($_.Exception.Message)"
        }
    }
    
    # Add variables
    foreach ($VarName in $Variables.Keys) {
        try {
            $variableEntry = New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry(
                $VarName,
                $Variables[$VarName],
                "Variable added by New-RunspaceSessionState"
            )
            [void]$initialSessionState.Variables.Add($variableEntry)
            Write-Verbose "Added variable: $VarName"
        }
        catch {
            Write-Warning "Failed to add variable '$VarName': $($_.Exception.Message)"
        }
    }
    
    return $initialSessionState
}

function New-RunspacePool {
    <#
    .SYNOPSIS
    Creates a new RunspacePool with specified configuration.
    
    .PARAMETER MinRunspaces
    Minimum number of runspaces in the pool
    
    .PARAMETER MaxRunspaces  
    Maximum number of runspaces in the pool (throttle limit)
    
    .PARAMETER SessionState
    InitialSessionState object (created by New-RunspaceSessionState)
    
    .PARAMETER Functions
    Array of function names to import (if SessionState not provided)
    
    .PARAMETER Modules
    Array of module names to import (if SessionState not provided)
    
    .PARAMETER Variables
    Hashtable of variables to add (if SessionState not provided)
    #>
    [CmdletBinding()]
    param(
        [int]$MinRunspaces = 1,
        [int]$MaxRunspaces = 3,
        [System.Management.Automation.Runspaces.InitialSessionState]$SessionState,
        [System.Collections.Generic.List[string]]$Functions = [System.Collections.Generic.List[string]]::new(),
        [System.Collections.Generic.List[string]]$Modules = [System.Collections.Generic.List[string]]::new(),
        [hashtable]$Variables = @{}
    )
    
    # Create session state if not provided
    if (-not $SessionState) {
        Write-Verbose "Creating new session state"
        $SessionState = New-RunspaceSessionState -Functions $Functions -Modules $Modules -Variables $Variables
    }
    
    # Create and open runspace pool
    try {
        $runspacePool = [runspacefactory]::CreateRunspacePool(
            $MinRunspaces,
            $MaxRunspaces, 
            $SessionState,
            $Host
        )
        [void]$runspacePool.Open()
        
        Write-Verbose "Created RunspacePool with $MinRunspaces-$MaxRunspaces runspaces"
        return $runspacePool
    }
    catch {
        Write-Error "Failed to create RunspacePool: $($_.Exception.Message)"
        return $null
    }
}

function New-RunspaceTask {
    <#
    .SYNOPSIS
    Creates a new runspace task with the specified script block.
    
    .PARAMETER RunspacePool
    The runspace pool to use
    
    .PARAMETER ScriptBlock
    The script block to execute
    
    .PARAMETER Parameters
    Array of parameters to pass to the script block
    
    .PARAMETER RunspaceId
    Unique identifier for this runspace task
    
    .PARAMETER TaskDescription
    Meaningful description of what this task does (e.g., "Processing SERVER01", "Backup Database_Prod")
    
    .PARAMETER TimeoutSeconds
    Timeout in seconds for this task
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.Runspaces.RunspacePool]$RunspacePool,
        
        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock,
        
        [System.Collections.Generic.List[object]]$Parameters = [System.Collections.Generic.List[object]]::new(),
        
        [string]$RunspaceId,
        
        [string]$TaskDescription,
        
        [int]$TimeoutSeconds = 30
    )
    
    # Handle auto-naming if no RunspaceId provided
    if (-not $RunspaceId) {
        $timestamp = Get-Date -Format "HHmmss"
        $milliseconds = (Get-Date).Millisecond.ToString("000")
        $RunspaceId = "Task_$timestamp$milliseconds"
    }
    
    # If no TaskDescription provided, use the RunspaceId as the description
    if (-not $TaskDescription) {
        $TaskDescription = $RunspaceId
    }
    
    try {
        # Create PowerShell instance
        $powerShell = [PowerShell]::Create()
        $powerShell.RunspacePool = $RunspacePool
        
        # Add script block with internal scope to avoid data "bleed"
        [void]$powerShell.AddScript($ScriptBlock, $true)
        
        # Add parameters
        foreach ($param in $Parameters) {
            [void]$powerShell.AddArgument($param)
        }
        
        # Start execution
        $asyncHandle = $powerShell.BeginInvoke()
        
        # Return task object
        return [PSCustomObject]@{
            RunspaceId      = $RunspaceId
            TaskDescription = $TaskDescription
            CurrentActivity = "Starting..."
            PowerShell      = $powerShell
            AsyncHandle     = $asyncHandle
            StartTime       = Get-Date
            TimeoutSeconds  = $TimeoutSeconds
            Status          = "Running"
            Results         = $null
            HasErrors       = $false
            HasWarnings     = $false
            Progress        = 0
        }
    }
    catch {
        Write-Error "Failed to create runspace task: $($_.Exception.Message)"
        return $null
    }
}

function Show-VisualProgress {
    <#
    .SYNOPSIS
    Helper function to display the visual progress screen consistently.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Collections.Generic.List[pscustomobject]]$Tasks,
        
        [Parameter(Mandatory)]
        [hashtable]$DisplayConfig,
        
        [int]$TotalTasks,
        
        [int]$PollingIntervalMs,
        
        [System.Collections.Generic.List[string]]$ActivityLog = [System.Collections.Generic.List[string]]::new(),
        
        [switch]$IsFinalDisplay
    )
    
    # Clear the screen for a clean display
    Clear-Host
    
    # Header information
    $headerText = if ($IsFinalDisplay) {
        "Monitoring $TotalTasks tasks - ALL COMPLETED"
    }
    else {
        "Monitoring $TotalTasks tasks with $PollingIntervalMs ms polling interval"
    }
    
    $headerColor = if ($IsFinalDisplay) { "Green" } else { "Gray" }
    
    Write-Host "Runspace Management - Started: $($Tasks[0].StartTime.ToString('HH:mm:ss'))" -ForegroundColor Gray
    Write-Host "$headerText`n" -ForegroundColor $headerColor
    
    # Progress header
    $progressHeaderColor = if ($IsFinalDisplay) { "Green" } else { "Magenta" }
    $progressHeaderText = if ($IsFinalDisplay) { "FINAL RESULTS" } else { "RUNSPACE PROGRESS" }
    
    Write-Host "$("="*70)" -ForegroundColor $progressHeaderColor
    Write-Host "$progressHeaderText - $(Get-Date -Format 'HH:mm:ss')" -ForegroundColor $progressHeaderColor
    Write-Host "$("="*70)" -ForegroundColor $progressHeaderColor
    
    # Sort tasks for consistent display
    $sortedTasks = $Tasks | Sort-Object { 
        if ($_.RunspaceId -match "Task(\d+)") { 
            [int]$matches[1] 
        }
        else { 
            $_.StartTime 
        } 
    }
    
    # Calculate padding for alignment
    $maxDescriptionLength = ($sortedTasks | ForEach-Object { $_.TaskDescription.Length } | Measure-Object -Maximum).Maximum
    
    # Show individual task progress/status
    foreach ($task in $sortedTasks) {
        $runtime = if ($task.Status -eq "Running") { 
            [math]::Round(((Get-Date) - $task.StartTime).TotalSeconds, 1)
        }
        else { 
            [math]::Round(((Get-Date) - $task.StartTime).TotalSeconds, 1)
        }
        
        # Progress calculation
        if ($IsFinalDisplay) {
            # For final display, show 100% for completed, 0% for failed
            $progress = if ($task.Status -eq "Completed") { 100 } else { 0 }
            $progressBar = if ($task.Status -eq "Completed") {
                $DisplayConfig.Console.ProgressChars.Filled * 10
            }
            else {
                $DisplayConfig.Console.ProgressChars.Empty * 10
            }
        }
        else {
            # For progress display, calculate based on runtime
            $progress = if ($task.Status -eq "Running") {
                [math]::Min([math]::Round(($runtime / $task.TimeoutSeconds) * 100), 100)
            }
            elseif ($task.Status -eq "Completed") {
                100
            }
            else {
                0
            }
            
            # Create progress bar
            $filledBars = [math]::Floor($progress / 10)
            $emptyBars = 10 - $filledBars
            $progressBar = $DisplayConfig.Console.ProgressChars.Filled * $filledBars + $DisplayConfig.Console.ProgressChars.Empty * $emptyBars
        }
        
        # Status display
        $statusIcon = $DisplayConfig.Console.Symbols[$task.Status]
        $statusColor = $DisplayConfig.Console.Colors[$task.Status]
        
        # Step description
        $stepDescription = switch ($task.Status) {
            "Running" { "Processing ($($runtime)s elapsed)" }
            "Completed" { "Completed in $($runtime)s" }
            "TimedOut" { "Timed out after $($runtime)s" }
            "Failed" { "Failed after $($runtime)s" }
            default { "Unknown status" }
        }
        
        # Pad description for alignment
        $paddedDescription = $task.TaskDescription.PadRight($maxDescriptionLength)
        
        Write-Host "$statusIcon $paddedDescription [$progressBar] $progress% - $stepDescription" -ForegroundColor $statusColor
    }
    
    # Summary stats
    $completedTasks = $Tasks | Where-Object { $_.Status -ne "Running" }
    $successfulTasks = $Tasks | Where-Object { $_.Status -eq "Completed" }  
    $failedTasks = $Tasks | Where-Object { $_.Status -in @("TimedOut", "Failed") }
    $runningTasks = $Tasks | Where-Object { $_.Status -eq "Running" }
    
    $completedSymbol = $DisplayConfig.Console.Symbols["Completed"]
    $failedSymbol = $DisplayConfig.Console.Symbols["Failed"] 
    $runningSymbol = $DisplayConfig.Console.Symbols["Running"]
    
    $summaryHeaderColor = if ($IsFinalDisplay) { "Green" } else { "Magenta" }
    $summaryHeaderText = if ($IsFinalDisplay) { "FINAL SUMMARY" } else { "SUMMARY" }
    
    Write-Host "`n$("="*70)" -ForegroundColor $summaryHeaderColor
    Write-Host "$summaryHeaderText`: $($completedTasks.Count)/$TotalTasks completed | $completedSymbol $($successfulTasks.Count) successful | $failedSymbol $($failedTasks.Count) failed | $runningSymbol $($runningTasks.Count) running" -ForegroundColor Cyan
    Write-Host "$("="*70)" -ForegroundColor $summaryHeaderColor
    
    # Show activity log (last few events)
    if ($ActivityLog.Count -gt 0) {
        Write-Host "`nRecent Activity:" -ForegroundColor Yellow
        $recentEvents = $ActivityLog | Select-Object -Last 10
        foreach ($ev in $recentEvents) {
            Write-Host "  $ev" -ForegroundColor Gray
        }
    }
    
    if ($IsFinalDisplay) {
        Write-Host ""  # Add spacing before next output
    }
}

function Wait-RunspaceTask {
    <#
    .SYNOPSIS
    Waits for runspace tasks to complete with timeout and progress monitoring.
    
    .PARAMETER Tasks
    Array of runspace task objects from New-RunspaceTask
    
    .PARAMETER PollingIntervalMs
    How often to check for completion (milliseconds)
    
    .PARAMETER OutputType
    Type of progress output to display
    
    .PARAMETER ProgressCallback
    Optional script block to call for custom progress handling
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Collections.Generic.List[pscustomobject]]$Tasks,
        
        [int]$PollingIntervalMs = 1000,
        
        [ValidateSet('Quiet', 'Basic', 'Visual', 'HtmlDashboard')]
        [string]$OutputType = 'Basic',
        
        [scriptblock]$ProgressCallback
    )
    
    # Define display configuration with console-safe characters
    $DisplayConfig = @{
        Console = @{
            Symbols       = @{
                "Running"   = "[*]"
                "Completed" = "[/]"
                "TimedOut"  = "[!]"
                "Failed"    = "[X]"
                "Unknown"   = "[?]"
            }
            Colors        = @{
                "Running"   = "Yellow"
                "Completed" = "Green"
                "TimedOut"  = "Red"
                "Failed"    = "Red"
                "Unknown"   = "Gray"
            }
            ProgressChars = @{
                "Filled" = "▓"
                "Empty"  = "░"
            }
        }
        Web     = @{
            Symbols       = @{
                "Running"   = "⚡"
                "Completed" = "✅"
                "TimedOut"  = "⏰"
                "Failed"    = "❌"
                "Unknown"   = "❓"
            }
            ProgressChars = @{
                "Filled" = "▓"
                "Empty"  = "░"
            }
        }
    }

    if ($OutputType -eq 'HtmlDashboard') {
        Write-Warning @"
HtmlDashboard mode is experimental. 
- Dashboard will launch in your browser
- You may need to manually close the browser tab when done
- Dashboard files will remain in default or specified Dashboard directory
- Use Ctrl+C to stop monitoring if needed
"@

        $response = Read-Host "Continue with WebDashboard mode? (y/N)"
        if ($response -ne 'y' -and $response -ne 'Y') {
            Write-Host "Switching to Visual mode instead..." -ForegroundColor Yellow
            $OutputType = 'Visual'
        }
    }

    $completedCount = 0
    $totalTasks = $Tasks.Count
    $lastVisualUpdate = Get-Date
    
    # Initialize activity log for Visual mode
    if ($null -eq $script:ActivityLog -or $script:ActivityLog -isnot [System.Collections.Generic.List[String]]) {
        $script:ActivityLog = [System.Collections.Generic.List[String]]::new()
    }
    
    Write-Verbose "Monitoring $totalTasks runspace tasks with '$OutputType' output"
    
    while ($completedCount -lt $totalTasks) {
        Start-Sleep -Milliseconds $PollingIntervalMs
        
        foreach ($task in ($Tasks | Where-Object { $_.Status -eq "Running" })) {
            $runtime = (Get-Date) - $task.StartTime
            $runtimeSeconds = [math]::Round($runtime.TotalSeconds, 1)
            
            # Check for timeout
            if ($runtimeSeconds -ge $task.TimeoutSeconds) {
                $symbol = $DisplayConfig.Console.Symbols["TimedOut"]
                if ($OutputType -ne 'Quiet') {
                    Write-Warning "$symbol Task $($task.TaskDescription) timed out after $runtimeSeconds seconds"
                }
                
                # Add to activity log
                [void]$script:ActivityLog.Add("$(Get-Date -Format 'HH:mm:ss') - $symbol $($task.TaskDescription) timed out after $runtimeSeconds seconds")
                
                try {
                    [void]$task.PowerShell.Stop()
                    [void]$task.PowerShell.Dispose()
                }
                catch {
                    if ($OutputType -ne 'Quiet') {
                        Write-Warning "Error stopping task $($task.TaskDescription): $($_.Exception.Message)"
                    }
                }
                
                $task.Status = "TimedOut"
                $task.Results = [pscustomobject]@{
                    Status         = "TimedOut"
                    RuntimeSeconds = $runtimeSeconds
                    TimeoutSeconds = $task.TimeoutSeconds
                }
                $completedCount++
            }
            # Check for completion
            elseif ($task.AsyncHandle.IsCompleted) {
                $symbol = $DisplayConfig.Console.Symbols["Completed"]
                $color = $DisplayConfig.Console.Colors["Completed"]
                
                if ($OutputType -eq 'Basic') {
                    Write-Host "$symbol Task $($task.TaskDescription) completed after $runtimeSeconds seconds" -ForegroundColor $color
                }
                
                # Add to activity log
                [void]$script:ActivityLog.Add("$(Get-Date -Format 'HH:mm:ss') - $symbol $($task.TaskDescription) completed after $runtimeSeconds seconds")
                
                try {
                    $task.Results = [pscustomobject]($task.PowerShell.EndInvoke($task.AsyncHandle))
                    [void]$task.PowerShell.Dispose()
                    $task.Status = "Completed"
                }
                catch {
                    if ($OutputType -ne 'Quiet') {
                        Write-Warning "Error getting results from task $($task.TaskDescription): $($_.Exception.Message)"
                    }
                    $task.Status = "Failed"
                    $task.HasErrors = $true
                    $task.Results = [pscustomobject]@{
                        Status = "Failed"
                        Error  = $_.Exception.Message
                    }
                }
                
                $completedCount++
            }
        }
        
        # Show progress based on OutputType
        if ($OutputType -eq 'Quiet') {
            # Show nothing
        }
        elseif ($OutputType -eq 'Basic') {
            $runningTasks = $Tasks | Where-Object { $_.Status -eq "Running" }
            Write-Host "Progress: $completedCount/$totalTasks completed, $($runningTasks.Count) running" -ForegroundColor Cyan
        }
        elseif ($OutputType -eq 'Visual') {
            # Show visual progress bars every few seconds
            $now = Get-Date
            if (($now - $lastVisualUpdate).TotalSeconds -ge 3) {
                $null = Show-VisualProgress -Tasks $Tasks -DisplayConfig $DisplayConfig -TotalTasks $totalTasks -PollingIntervalMs $PollingIntervalMs -ActivityLog $script:ActivityLog
                $lastVisualUpdate = $now
            }
        }
        elseif ($OutputType -eq 'HtmlDashboard') {
            $null = Export-RunspaceHtmlDashboard -Tasks $Tasks -LaunchBrowser

            if (-not $dashboardLaunched) {
                $fullPath = Resolve-Path "Dashboard\dashboard.html"
                $null = Start-Process $fullPath
                Write-Host "Dashboard launched in browser!" -ForegroundColor Cyan
                $dashboardLaunched = $true
            }
        }
        
        # Call progress callback if provided
        if ($ProgressCallback) {
            $progressInfo = @{
                CompletedCount = $completedCount
                TotalTasks     = $totalTasks
                RunningTasks   = $Tasks | Where-Object { $_.Status -eq "Running" }
                CompletedTasks = $Tasks | Where-Object { $_.Status -ne "Running" }
            }
            & $ProgressCallback $progressInfo
        }
    }
    
    # finalizations
    if ($OutputType -ne 'Quiet') {
        Write-Verbose "All tasks completed"
    }
    
    if ($OutputType -eq 'Visual') {
        # Force one final display update to show completion
        # Force one final display update to show completion
        Start-Sleep -Milliseconds 500  # Brief pause to let everything settle
        $null = Show-VisualProgress -Tasks $Tasks -DisplayConfig $DisplayConfig -TotalTasks $totalTasks -PollingIntervalMs $PollingIntervalMs -ActivityLog $script:ActivityLog -IsFinalDisplay
    }
}

function Get-RunspaceResults {
    <#
    .SYNOPSIS
    Extracts and formats results from completed runspace tasks.
    
    .PARAMETER Tasks
    Array of completed runspace task objects
    
    .PARAMETER IncludeMetadata
    Whether to include timing and status metadata
    
    .PARAMETER ExportPath
    Optional path to export results to CSV/XML
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Collections.Generic.List[pscustomobject]]$Tasks,
        
        [switch]$IncludeMetadata,
        
        [string]$ExportPath
    )
    
    $results = [System.Collections.Generic.List[object]]::new()
    
    foreach ($task in $Tasks) {
        $result = [PSCustomObject]@{
            RunspaceId      = $task.RunspaceId
            TaskDescription = $task.TaskDescription
            Status          = $task.Status
            StartTime       = $task.StartTime
            RuntimeSeconds  = if ($task.Status -eq "Running") { 
                [math]::Round(((Get-Date) - $task.StartTime).TotalSeconds, 1) 
            }
            else { 
                [math]::Round(((Get-Date) - $task.StartTime).TotalSeconds, 1) 
            }
            TimeoutSeconds  = $task.TimeoutSeconds
            HasErrors       = $task.HasErrors
            HasWarnings     = $task.HasWarnings
            Results         = $task.Results
        }
        
        if (-not $IncludeMetadata) {
            $result = $result | Select-Object RunspaceId, TaskDescription, Status, Results
        }
        
        [void]$results.Add($result)
    }
    
    # Export if path provided
    if ($ExportPath) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        
        if ($ExportPath.EndsWith('.xml')) {
            $null = $results | Export-Clixml -Path $ExportPath -Force
            Write-Verbose "Results exported to XML: $ExportPath"
        }
        elseif ($ExportPath.EndsWith('.csv')) {
            $null = $results | Export-Csv -Path $ExportPath -NoTypeInformation -Force
            Write-Verbose "Results exported to CSV: $ExportPath"
        }
        else {
            # Default to XML
            $xmlPath = "$ExportPath`_$timestamp.xml"
            $null = $results | Export-Clixml -Path $xmlPath -Force
            Write-Verbose "Results exported to XML: $xmlPath"
        }
    }
    
    return $results
}

function Export-RunspaceHtmlDashboard {
    <#
    .SYNOPSIS
    Exports runspace progress to a beautiful HTML dashboard with embedded JSON data.
    
    .PARAMETER Tasks
    Array of runspace task objects
    
    .PARAMETER OutputPath
    Directory path for dashboard files
    
    .PARAMETER RefreshIntervalSeconds
    How often the dashboard refreshes
    
    .PARAMETER LaunchBrowser
    Whether to launch the dashboard in browser (only on first creation)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Collections.Generic.List[pscustomobject]]$Tasks,
        
        [string]$OutputPath = "Dashboard",
        
        [int]$RefreshIntervalSeconds = 2,
        
        [switch]$LaunchBrowser
    )
    
    # Create output directory
    if (-not (Test-Path $OutputPath)) {
        $null = New-Item -ItemType Directory -Path $OutputPath -Force
    }
    
    # Calculate status data
    $statusData = @{
        LastUpdate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Tasks      = $Tasks | ForEach-Object {
            $runtime = ((Get-Date) - $_.StartTime).TotalSeconds
            $progress = if ($_.Status -eq "Running") { 
                [math]::Min([math]::Round(($runtime / $_.TimeoutSeconds) * 100), 100)
            }
            elseif ($_.Status -eq "Completed") { 
                100 
            }
            else { 
                0 
            }
            
            @{
                Id             = $_.RunspaceId
                Description    = $_.TaskDescription
                Status         = $_.Status
                Progress       = $progress
                Runtime        = [math]::Round($runtime, 1)
                TimeoutSeconds = $_.TimeoutSeconds
            }
        }
        Summary    = @{
            Total     = $Tasks.Count
            Completed = ($Tasks | Where-Object { $_.Status -eq "Completed" }).Count
            Running   = ($Tasks | Where-Object { $_.Status -eq "Running" }).Count
            Failed    = ($Tasks | Where-Object { $_.Status -in @("TimedOut", "Failed") }).Count
        }
    }
    
    # Convert to JSON for embedding
    $jsonData = $statusData | ConvertTo-Json -Depth 3 -Compress
    
    # Create HTML with embedded JSON data
    $htmlPath = "$OutputPath\dashboard.html"
    $dashboardCreated = $false

    if (-not (Test-Path $htmlPath)) {
        $dashboardCreated = $true
        Write-Verbose "Dashboard updated with embedded data at: $htmlPath"
    }

    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>PowerShell Runspace Dashboard</title>
    <meta charset="utf-8">
    <meta http-equiv="refresh" content="$RefreshIntervalSeconds">
    <style>
        body { 
            font-family: 'Segoe UI', 'SF Pro Display', -apple-system, BlinkMacSystemFont, Arial, sans-serif; 
            margin: 0; 
            padding: 20px; 
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            color: #333;
        }
        
        .container { 
            max-width: 1200px; 
            margin: 0 auto; 
            background: rgba(255,255,255,0.95); 
            border-radius: 15px; 
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
            overflow: hidden;
        }
        
        .header { 
            background: linear-gradient(135deg, #2c3e50 0%, #34495e 100%); 
            color: white; 
            padding: 30px; 
            text-align: center;
            border-bottom: 4px solid #3498db;
        }
        
        .header h1 { 
            margin: 0 0 10px 0; 
            font-size: 2.5em; 
            font-weight: 300; 
            letter-spacing: 1px;
        }
        
        .header p { 
            margin: 0; 
            opacity: 0.9; 
            font-size: 1.1em; 
        }
        
        .summary { 
            display: grid; 
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); 
            gap: 20px; 
            padding: 30px; 
            background: #f8f9fa;
        }
        
        .summary-card { 
            background: white; 
            padding: 25px; 
            border-radius: 10px; 
            box-shadow: 0 4px 6px rgba(0,0,0,0.07);
            text-align: center; 
            border-left: 4px solid;
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }
        
        .summary-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 15px rgba(0,0,0,0.1);
        }
        
        .summary-card.total { border-left-color: #3498db; }
        .summary-card.completed { border-left-color: #27ae60; }
        .summary-card.running { border-left-color: #f39c12; }
        .summary-card.failed { border-left-color: #e74c3c; }
        
        .summary-card h3 { 
            margin: 0 0 10px 0; 
            font-size: 2.5em; 
            font-weight: 600; 
        }
        
        .summary-card p { 
            margin: 0; 
            color: #666; 
            font-size: 0.9em;
            text-transform: uppercase;
            letter-spacing: 1px;
        }
        
        .task-container { 
            padding: 30px; 
        }
        
        .task-container h2 { 
            margin: 0 0 25px 0; 
            color: #2c3e50; 
            font-size: 1.8em;
            font-weight: 300;
            border-bottom: 2px solid #ecf0f1;
            padding-bottom: 10px;
        }
        
        .task { 
            margin-bottom: 20px; 
            background: white;
            border-radius: 8px;
            padding: 20px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
            border-left: 4px solid #ecf0f1;
            transition: all 0.3s ease;
        }
        
        .task.status-running { border-left-color: #f39c12; }
        .task.status-completed { border-left-color: #27ae60; }
        .task.status-failed { border-left-color: #e74c3c; }
        .task.status-timedout { border-left-color: #e67e22; }
        
        .task-header { 
            display: flex; 
            justify-content: space-between; 
            align-items: center; 
            margin-bottom: 15px; 
        }
        
        .task-name {
            display: flex;
            align-items: center;
            font-weight: 600;
            font-size: 1.1em;
        }
        
        .status-icon { 
            font-size: 20px; 
            margin-right: 12px;
            min-width: 20px;
        }
        
        .task-status {
            font-size: 0.9em;
            color: #666;
            background: #f8f9fa;
            padding: 4px 12px;
            border-radius: 20px;
        }
        
        .progress-container {
            margin-top: 10px;
        }
        
        .progress-label {
            display: flex;
            justify-content: space-between;
            font-size: 0.85em;
            color: #666;
            margin-bottom: 5px;
        }
        
        .progress-bar { 
            width: 100%; 
            height: 8px; 
            background-color: #ecf0f1; 
            border-radius: 10px; 
            overflow: hidden;
            box-shadow: inset 0 1px 3px rgba(0,0,0,0.1);
        }
        
        .progress-fill { 
            height: 100%; 
            transition: width 0.6s ease; 
            border-radius: 10px;
            position: relative;
            overflow: hidden;
        }
        
        .progress-fill::after {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            bottom: 0;
            right: 0;
            background-image: linear-gradient(45deg, rgba(255,255,255,0.2) 25%, transparent 25%, transparent 50%, rgba(255,255,255,0.2) 50%, rgba(255,255,255,0.2) 75%, transparent 75%, transparent);
            background-size: 20px 20px;
            animation: move-stripes 1s linear infinite;
        }
        
        @keyframes move-stripes {
            0% { background-position: 0 0; }
            100% { background-position: 20px 20px; }
        }
        
        .status-running .progress-fill { 
            background: linear-gradient(135deg, #f39c12, #e67e22); 
        }
        .status-completed .progress-fill { 
            background: linear-gradient(135deg, #27ae60, #219a52); 
        }
        .status-failed .progress-fill { 
            background: linear-gradient(135deg, #e74c3c, #c0392b); 
        }
        .status-timedout .progress-fill { 
            background: linear-gradient(135deg, #e67e22, #d35400); 
        }
        
        .last-update { 
            color: #7f8c8d; 
            font-size: 0.85em; 
            text-align: center; 
            margin-top: 30px; 
            padding: 20px;
            background: #f8f9fa;
            border-radius: 0 0 15px 15px;
        }
        
        .pulse {
            animation: pulse 2s infinite;
        }
        
        @keyframes pulse {
            0% { opacity: 1; }
            50% { opacity: 0.7; }
            100% { opacity: 1; }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🚀 PowerShell Runspace Dashboard</h1>
            <p>Real-time monitoring of parallel runspace execution</p>
        </div>
        
        <div id="content">
            <p style="text-align: center; padding: 40px; color: #666;">Loading dashboard data...</p>
        </div>
    </div>
    
    <script>
        // Embedded JSON data
        const dashboardData = $jsonData;
        
        function updateDashboard(data) {
            const summary = data.Summary;
            const tasks = data.Tasks;
            
            let html = '';
            
            // Summary cards
            html += '<div class="summary">';
            html += '<div class="summary-card total"><h3>' + summary.Total + '</h3><p>Total Tasks</p></div>';
            html += '<div class="summary-card completed"><h3 style="color: #27ae60;">' + summary.Completed + '</h3><p>Completed</p></div>';
            html += '<div class="summary-card running"><h3 style="color: #f39c12;">' + summary.Running + '</h3><p>Running</p></div>';
            html += '<div class="summary-card failed"><h3 style="color: #e74c3c;">' + summary.Failed + '</h3><p>Failed</p></div>';
            html += '</div>';
            
            // Task details
            html += '<div class="task-container">';
            html += '<h2>📊 Task Progress</h2>';
            
            tasks.forEach(task => {
                const statusClass = 'status-' + task.Status.toLowerCase();
                const statusIcon = task.Status === 'Running' ? '⚡' : 
                                 task.Status === 'Completed' ? '✅' : 
                                 task.Status === 'TimedOut' ? '⏰' : '❌';
                                 
                const pulseClass = task.Status === 'Running' ? 'pulse' : '';
                
                html += '<div class="task ' + statusClass + '">';
                html += '<div class="task-header">';
                html += '<div class="task-name"><span class="status-icon ' + pulseClass + '">' + statusIcon + '</span>' + task.Description + '</div>';
                html += '<div class="task-status">' + task.Status + ' (' + task.Runtime + 's)</div>';
                html += '</div>';
                
                html += '<div class="progress-container">';
                html += '<div class="progress-label">';
                html += '<span>Progress</span>';
                html += '<span>' + task.Progress + '%</span>';
                html += '</div>';
                html += '<div class="progress-bar">';
                html += '<div class="progress-fill" style="width: ' + task.Progress + '%"></div>';
                html += '</div>';
                html += '</div>';
                
                html += '</div>';
            });
            
            html += '</div>';
            html += '<div class="last-update">⏱️ Last updated: ' + data.LastUpdate + '</div>';
            
            document.getElementById('content').innerHTML = html;
        }
        
        // Load embedded data
        updateDashboard(dashboardData);
    </script>
</body>
</html>
"@
    
    $html | Out-File $htmlPath -Force

    # Only launch browser if dashboard was just created AND LaunchBrowser was requested
    if ($LaunchBrowser -and $dashboardCreated) {
        $fullPath = Resolve-Path $htmlPath
        $null = Start-Process $fullPath
        Write-Host "Dashboard launched in browser..." -ForegroundColor Cyan
    }
    
    Write-Host "Dashboard exported to: $htmlPath" -ForegroundColor Green
}

function Stop-RunspacePool {
    <#
    .SYNOPSIS
    Properly closes and disposes of a runspace pool.
    
    .PARAMETER RunspacePool
    The runspace pool to close
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.Runspaces.RunspacePool]$RunspacePool
    )
    
    try {
        if ($RunspacePool.RunspacePoolStateInfo.State -eq 'Opened') {
            [void]$RunspacePool.Close()
            Write-Verbose "RunspacePool closed"
        }
        
        [void]$RunspacePool.Dispose()
        Write-Verbose "RunspacePool disposed"
    }
    catch {
        Write-Warning "Error closing RunspacePool: $($_.Exception.Message)"
    }
}

#endregion

Write-Verbose "Runspace Management Functions Loaded!"

<# Example usage:

Invoke-ParallelRunspaceExample -NumberOfTasks 8 -MaxConcurrentRunspaces 4 -OutputType 'Visual'

#>
