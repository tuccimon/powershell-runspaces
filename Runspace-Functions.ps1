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
    $failedModules = @()
    foreach ($ModuleName in ($Modules | Sort-Object -Unique)) {
        try {
            [void]$initialSessionState.ImportPSModule($ModuleName)
            # Verify module exists in session state
            $moduleExists = $null = $initialSessionState.Commands | Where-Object { 
                $_.Name -eq $ModuleName -or $_.Module -eq $ModuleName 
            }
            if (-not $moduleExists) {
                $failedModules += $ModuleName
                Write-Warning "Module '$ModuleName' may not have imported correctly"
            }
        }
        catch {
            $failedModules += $ModuleName
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
        $runspacePool.ApartmentState = 'STA'

        # Attach tracker to the pool object itself (extend it)
        $null = $runspacePool | Add-Member -MemberType NoteProperty -Name "Tracker" -Value ([hashtable]::Synchronized(@{})) -Force

        [void]$runspacePool.Open()

        Write-Verbose "Created RunspacePool with $MinRunspaces-$MaxRunspaces runspaces"
        return $runspacePool
    }
    catch {
        Write-Error "Failed to create RunspacePool: $($_.Exception.Message)"
        return $null
    }
}


function Import-RunspaceScriptBlock {
    <#
    .SYNOPSIS
    Imports a script block, conforms it for runspace, and then outputs the updated script block
   
    .PARAMETER ScriptBlock
    The script block to import
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock
    )

    ### until a good solution is found, we will be blocking
    if ($ScriptBlock.Ast.UsingStatements) {
        throw "Using statements are not supported in runspace script blocks"
    }

    if ($ScriptBlock.Ast.ParamBlock -and $ScriptBlock.Ast.ParamBlock.Attributes) {
        throw "Parameter attributes (like [CmdletBinding()]) are not supported"
    }
    ###########################################################################

    # Create a new ArrayList for parameters
    $parameters = [System.Collections.ArrayList]::new()

    # Extract existing parameters
    if ($ScriptBlock.Ast.ParamBlock) {
        foreach ($param in $ScriptBlock.Ast.ParamBlock.Parameters.Extent.Text) {
            [void]$parameters.Add($param)
        }
    }

    # inject runspace params
    [void]$parameters.Add('$Tracker')
    [void]$parameters.Add('$TaskGuid')
    $ParamLine = "param({0})" -f ($parameters -join ', ')

    # create signal
    $signalLine = '$Tracker[$TaskGuid] = @{ ActualStartTime = Get-Date }'

    # Get the body content properly
    $scriptText = $ScriptBlock.Ast.Extent.Text
    $paramText = $ScriptBlock.Ast.ParamBlock.Extent.Text

    if ([string]::IsNullOrWhiteSpace($paramText)) {
        $bodyContent = "$scriptText".Trim('{}').Trim()
    }
    else {
        $bodyContent = "$scriptText".Replace($paramText, '').Trim('{}').Trim()
    }

    # build updated script block text
    $updated = @"
$ParamLine

$signalLine

$bodyContent
"@

    return ([scriptblock]::Create($updated))
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
    
    # create unique identifier for task
    $taskGuid = [guid]::NewGuid().ToString()

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

        # transform script block for runspace purposes
        $updatedScriptBlock = Import-RunspaceScriptBlock -ScriptBlock $ScriptBlock

        # Add script block with internal scope to avoid data "bleed"
        [void]$powerShell.AddScript($updatedScriptBlock, $true)

        # add tracker and taskGuid to parameters
        [void]$Parameters.Add($RunspacePool.Tracker)
        [void]$Parameters.Add($taskGuid)

        # Add parameters
        foreach ($param in $Parameters) {
            [void]$powerShell.AddArgument($param)
        }

        # Start execution
        $asyncHandle = $powerShell.BeginInvoke()
        
        # Return task object
        return [PSCustomObject]@{
            Guid            = $taskGuid
            RunspaceId      = $RunspaceId
            TaskDescription = $TaskDescription
            CurrentActivity = "Starting..."
            PowerShell      = $powerShell
            AsyncHandle     = $asyncHandle
            StartTime       = Get-Date
            ActualStartTime = $null
            Duration        = $null
            TimeoutSeconds  = $TimeoutSeconds
            Status          = "Queued"
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
        # In Show-VisualProgress, replace the runtime calculation
        $runtime = if ($task.Status -eq "Completed" -and $task.ActualStartTime) {
            $task.Duration
        }
        elseif ($task.Status -eq "Running" -and $task.ActualStartTime) {
            # Use ActualStartTime for running tasks
            [math]::Round(((Get-Date) - $task.ActualStartTime).TotalSeconds, 1)
        }
        else {
            # Fallback to StartTime
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
            $progress = if ($task.Status -eq "Running" -and $task.ActualStartTime) {
                $elapsed = ((Get-Date) - $task.ActualStartTime).TotalSeconds
                [math]::Min([math]::Round(($elapsed / $task.TimeoutSeconds) * 100), 100)
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
            "Queued" { "Task queued - not started" }
            default { "Unknown status" }
        }
        
        # Pad description for alignment
        $paddedDescription = $task.TaskDescription.PadRight($maxDescriptionLength)
        
        Write-Host "$statusIcon $paddedDescription [$progressBar] $progress% - $stepDescription" -ForegroundColor $statusColor
    }
    
    # Summary stats
    $completedTasks = $Tasks | Where-Object { $_.Status -in @("Completed", "TimedOut", "Failed") }
    $successfulTasks = $Tasks | Where-Object { $_.Status -eq "Completed" }  
    $failedTasks = $Tasks | Where-Object { $_.Status -in @("TimedOut", "Failed") }
    $runningTasks = $Tasks | Where-Object { $_.Status -eq "Running" }
    $notStartedTasks = $Tasks | Where-Object { $_.Status -eq "Queued" }
    
    $completedSymbol = $DisplayConfig.Console.Symbols["Completed"]
    $failedSymbol = $DisplayConfig.Console.Symbols["Failed"] 
    $runningSymbol = $DisplayConfig.Console.Symbols["Running"]
    $queuedSymbol = $DisplayConfig.Console.Symbols["Queued"]
    
    $summaryHeaderColor = if ($IsFinalDisplay) { "Green" } else { "Magenta" }
    $summaryHeaderText = if ($IsFinalDisplay) { "FINAL SUMMARY" } else { "SUMMARY" }
    
    Write-Host "`n$("="*70)" -ForegroundColor $summaryHeaderColor
    Write-Host "$summaryHeaderText`: $($completedTasks.Count)/$TotalTasks completed | $completedSymbol $($successfulTasks.Count) successful | $failedSymbol $($failedTasks.Count) failed | $queuedSymbol $($notStartedTasks.Count) queued | $runningSymbol $($runningTasks.Count) running" -ForegroundColor Cyan
    Write-Host "$("="*70)" -ForegroundColor $summaryHeaderColor

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
    
    .PARAMETER Force
    Optional switch, that applies to HtmlDashboard, that prevents the confirmation prompt.

    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Collections.Generic.List[pscustomobject]]$Tasks,
        
        [int]$PollingIntervalMs = 1000,
        
        [ValidateSet('Quiet', 'Basic', 'Visual', 'HtmlDashboard')]
        [string]$OutputType = 'Basic',
        
        [switch]$Force
    )
    
    # Reset dashboard state for new run
    if ($OutputType -eq 'HtmlDashboard') {
        $script:lastDashboardState = $null
        $script:dashboardLaunched = $false
        Write-Verbose "Reset dashboard state for new monitoring session"
    }

    # Define display configuration with console-safe characters
    $DisplayConfig = @{
        Console = @{
            Symbols       = @{
                "Queued"    = "[ ]"
                "Running"   = "[*]"
                "Completed" = "[/]"
                "TimedOut"  = "[!]"
                "Failed"    = "[X]"
                "Unknown"   = "[?]"
            }
            Colors        = @{
                "Queued"    = "Gray"
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
                "Queued"    = "⏳"      # Hourglass
                "Running"   = "⚡"      # Lightning
                "Completed" = "✅"      # Check mark
                "TimedOut"  = "⏰"      # Alarm clock
                "Failed"    = "❌"      # X
                "Unknown"   = "❓"      # Question mark
            }
            ProgressChars = @{
                "Filled" = "▓"
                "Empty"  = "░"
            }
        }
    }

    if ($OutputType -eq 'HtmlDashboard' -and !($Force)) {
        $confirmationMessage = @"
HtmlDashboard mode is experimental. 
- Dashboard will launch in your browser
- You may need to manually close the browser tab when done
- Dashboard files will remain in default or specified Dashboard directory
- Use Ctrl+C to stop monitoring if needed

Are you sure you want to use this output mode?
"@
        $choice = $host.UI.PromptForChoice("HTML Dashboard", $confirmationMessage, @('&Yes','&No'), 1)
        if ($choice -ne 0) {
            Write-Host "Switching to Visual Mode instead..." -ForegroundColor Yellow
            $OutputType = 'Visual'
        }
    }

    $completedCount = 0
    $totalTasks = $Tasks.Count
    $lastVisualUpdate = Get-Date
    
    Write-Verbose "Monitoring $totalTasks runspace tasks with '$OutputType' output"
    
    while ($completedCount -lt $totalTasks) {
        Start-Sleep -Milliseconds $PollingIntervalMs
        $now = Get-Date
    
        # Check tracker for tasks that have actually started
        foreach ($task in ($Tasks | Where-Object { $_.Status -eq "Queued" })) {
            $tracker = $task.PowerShell.RunspacePool.Tracker
            if ($tracker.ContainsKey($task.Guid)) {
                $task.Status = "Running"
                $task.ActualStartTime = $tracker[$task.Guid].ActualStartTime
            }
        }

        # Now process running tasks
        foreach ($task in ($Tasks | Where-Object { $_.Status -eq "Running" })) {
            # Use ActualStartTime for runtime calculation, fallback to StartTime if something's wrong
            $startTimeForTimeout = if ($task.ActualStartTime) { 
                $task.ActualStartTime 
            }
            else { 
                $task.StartTime 
            }
    
            $runtime = $now - $startTimeForTimeout
            $runtimeSeconds = [math]::Round($runtime.TotalSeconds, 1)
            $task.Duration = $runtimeSeconds
    
            # Set deadline based on actual start time
            $deadline = $startTimeForTimeout.AddSeconds($task.TimeoutSeconds)
    
            # Check timeout against actual start time
            if ($now -gt $deadline) {
                $symbol = $DisplayConfig.Console.Symbols["TimedOut"]
                if ($OutputType -eq 'Basic') {
                    Write-Warning "$symbol Task $($task.TaskDescription) timed out after $runtimeSeconds seconds (limit: $($task.TimeoutSeconds)s)"
                }
            
                try {
                    # Stop just this specific task
                    [void]$task.PowerShell.Stop()
                    [void]$task.PowerShell.Dispose()
                }
                catch {
                    if ($OutputType -notin @('Quiet', 'HtmlDashboard')) {
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
            
                if ($OutputType -eq 'Basic') {
                    Write-Host "Task $($task.TaskDescription) marked as TimedOut, completedCount=$completedCount" -ForegroundColor Yellow
                }
            }
            # Check for completion
            elseif ($task.AsyncHandle.IsCompleted) {
                $symbol = $DisplayConfig.Console.Symbols["Completed"]
                $color = $DisplayConfig.Console.Colors["Completed"]
            
                if ($OutputType -eq 'Basic') {
                    Write-Host "$symbol Task $($task.TaskDescription) completed after $runtimeSeconds seconds" -ForegroundColor $color
                }
            
                try {
                    $task.Results = [pscustomobject]($task.PowerShell.EndInvoke($task.AsyncHandle))
                    [void]$task.PowerShell.Dispose()
                    $task.Status = "Completed"
                }
                catch {
                    if ($OutputType -notin @('Quiet', 'HtmlDashboard')) {
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
                
                if ($OutputType -eq 'Basic') {
                    Write-Host "Task $($task.TaskDescription) completed, completedCount=$completedCount" -ForegroundColor Green
                }
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
                $null = Show-VisualProgress -Tasks $Tasks -DisplayConfig $DisplayConfig -TotalTasks $totalTasks -PollingIntervalMs $PollingIntervalMs
                $lastVisualUpdate = $now
            }
        }
        elseif ($OutputType -eq 'HtmlDashboard') {
            # Only export if tasks have changed OR it's the first run
            $currentState = ($Tasks | ForEach-Object { "$($_.RunspaceId):$($_.Status):$($_.Progress)" }) -join '|'
            if ($null -eq $script:lastDashboardState -or $script:lastDashboardState -ne $currentState) {
                # Pass -Quiet to suppress the "exported" message on subsequent updates
                $quiet = $script:dashboardLaunched -eq $true
                $null = Export-RunspaceHtmlDashboard -Tasks $Tasks -LaunchBrowser:(-not $script:dashboardLaunched) -Quiet:$quiet
                $script:lastDashboardState = $currentState
        
                if (-not $script:dashboardLaunched) {
                    $script:dashboardLaunched = $true
                }
            }
        }
    }
    
    # finalizations
    if ($OutputType -ne 'Quiet') {
        Write-Verbose "All tasks completed"
    }
    
    if ($OutputType -eq 'Visual') {
        # Force one final display update to show completion
        Start-Sleep -Milliseconds 500  # Brief pause to let everything settle
        $null = Show-VisualProgress -Tasks $Tasks -DisplayConfig $DisplayConfig -TotalTasks $totalTasks -PollingIntervalMs $PollingIntervalMs -IsFinalDisplay
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
            ActualStartTime = $task.ActualStartTime
            RuntimeSeconds  = $task.Duration || [math]::Round(((Get-Date) - $task.StartTime).TotalSeconds, 1)
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
    Exports runspace progress to a HTML dashboard with embedded JSON data.
    
    .PARAMETER Tasks
    Array of runspace task objects
    
    .PARAMETER OutputPath
    Directory path for dashboard files
    
    .PARAMETER RefreshIntervalSeconds
    How often the dashboard refreshes
    
    .PARAMETER LaunchBrowser
    Whether to launch the dashboard in browser (only on first creation)

    .PARAMETER Quiet
    Suppresses the "Dashboard exported" message
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Collections.Generic.List[pscustomobject]]$Tasks,
        
        [string]$OutputPath = (Join-Path -Path $PSScriptRoot -ChildPath "Dashboard"),
        
        [int]$RefreshIntervalSeconds = 2,
        
        [switch]$LaunchBrowser,

        [switch]$Quiet
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
            Queued    = ($Tasks | Where-Object { $_.Status -eq "Queued" }).Count
        }
    }
    
    # Convert to JSON for embedding
    $jsonData = $statusData | ConvertTo-Json -Depth 3 -Compress
    
    # Create HTML with embedded JSON data and UTF-8 encoding
    $htmlPath = "$OutputPath\dashboard.html"

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="refresh" content="$RefreshIntervalSeconds">
    <title>PowerShell Runspace Dashboard</title>
    <style>
        * { 
            box-sizing: border-box; 
            margin: 0;
            padding: 0;
        }
        
        body { 
            font-family: 'Segoe UI', 'SF Pro Display', -apple-system, BlinkMacSystemFont, 'Apple Color Emoji', 'Segoe UI Emoji', 'Segoe UI Symbol', 'Noto Color Emoji', Arial, sans-serif;
            margin: 0; 
            padding: 20px; 
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
            min-height: 100vh; 
            color: #333; 
        }
        
        .container { 
            max-width: 1600px;
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
            padding: 20px; 
        }
        
        .task-container h2 { 
            margin: 0 0 20px 0; 
            color: #2c3e50; 
            font-size: 1.6em; 
            font-weight: 300; 
            border-bottom: 1px solid #ecf0f1; 
            padding-bottom: 8px; 
        }
        
        /* Responsive grid - cards will expand to fill width */
        .task-grid { 
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
            gap: 20px;
            padding: 20px;
            width: 100%;
        }
        
        /* For very large screens, allow cards to grow */
        @media (min-width: 1920px) {
            .task-grid {
                grid-template-columns: repeat(auto-fill, minmax(350px, 1fr));
                gap: 25px;
            }
        }
        
        /* For smaller screens, make cards more compact */
        @media (max-width: 768px) {
            .task-grid {
                grid-template-columns: 1fr;
                gap: 15px;
                padding: 15px;
            }
            
            .task {
                min-height: 180px;
            }
        }
        
        /* Card styling */
        .task { 
            display: flex; 
            flex-direction: column; 
            min-height: 200px;
            background: white; 
            border-radius: 12px; 
            padding: 20px; 
            box-shadow: 0 4px 12px rgba(0,0,0,0.1); 
            border-left: 6px solid; 
            transition: all 0.2s ease;
            width: 100%;
        }
        
        .task.status-running { border-left-color: #f39c12; }
        .task.status-completed { border-left-color: #27ae60; }
        .task.status-failed { border-left-color: #e74c3c; }
        .task.status-timedout { border-left-color: #e67e22; }
        .task.status-queued { border-left-color: #95a5a6; }
        
        .task:hover { 
            transform: translateY(-2px); 
            box-shadow: 0 8px 20px rgba(0,0,0,0.15); 
        }
        
        .task-header { 
            flex: 1; 
            display: flex; 
            flex-direction: column; 
            gap: 12px; 
        }
        
        .task-name { 
            font-weight: 600; 
            font-size: 1rem; 
            line-height: 1.4; 
            word-break: break-word;
            display: flex;
            align-items: flex-start;
            gap: 8px;
        }
        
        .status-icon { 
            font-size: 20px; 
            flex-shrink: 0;
            display: inline-block;
        }
        
        .task-status { 
            font-size: 0.85em; 
            color: #666; 
            background: #f8f9fa; 
            padding: 4px 10px; 
            border-radius: 12px; 
            display: inline-block;
            width: fit-content;
        }
        
        .progress-container { 
            margin-top: auto; 
            padding-top: 15px; 
        }
        
        .progress-label { 
            display: flex; 
            justify-content: space-between; 
            font-size: 0.85em; 
            color: #666; 
            margin-bottom: 6px; 
        }
        
        .progress-bar { 
            width: 100%; 
            height: 8px; 
            background-color: #ecf0f1; 
            border-radius: 6px; 
            overflow: hidden; 
            box-shadow: inset 0 1px 2px rgba(0,0,0,0.05); 
        }
        
        .progress-fill { 
            height: 100%; 
            transition: width 0.5s ease; 
            border-radius: 6px; 
        }
        
        .pulse { 
            animation: pulse 1.5s infinite; 
        }
        
        @keyframes pulse { 
            0% { opacity: 1; } 
            50% { opacity: 0.6; } 
            100% { opacity: 1; } 
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
        
        .status-queued .progress-fill { 
            background: linear-gradient(135deg, #95a5a6, #7f8c8d); 
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
        
        /* Optional: Add a subtle animation for running tasks */
        @keyframes subtleGlow {
            0% {
                box-shadow: 0 4px 12px rgba(243, 156, 18, 0.1);
            }
            50% {
                box-shadow: 0 4px 12px rgba(243, 156, 18, 0.3);
            }
            100% {
                box-shadow: 0 4px 12px rgba(243, 156, 18, 0.1);
            }
        }
        
        .task.status-running {
            animation: subtleGlow 2s ease-in-out infinite;
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
        const dashboardData = $jsonData;
        
        function getStatusIcon(status) {
            const icons = {
                'Queued': '⏳',
                'Running': '⚡',
                'Completed': '✅',
                'TimedOut': '⏰',
                'Failed': '❌'
            };
            return icons[status] || '❓';
        }
        
        function updateDashboard(data) {
            const summary = data.Summary;
            const tasks = data.Tasks;
            let html = '';
            
            // Summary cards
            html += '<div class="summary">';
            html += '<div class="summary-card total"><h3>' + summary.Total + '</h3><p>Total Tasks</p></div>';
            html += '<div class="summary-card completed"><h3>' + summary.Completed + '</h3><p>Completed</p></div>';
            html += '<div class="summary-card running"><h3>' + summary.Running + '</h3><p>Running</p></div>';
            html += '<div class="summary-card failed"><h3>' + summary.Failed + '</h3><p>Failed</p></div>';
            html += '<div class="summary-card queued"><h3>' + summary.Queued + '</h3><p>Queued</p></div>';
            html += '</div>';
            
            // Task grid
            html += '<div class="task-container">';
            html += '<h2>📊 Task Progress</h2>';
            html += '<div class="task-grid">';
            
            tasks.forEach(task => {
                const statusClass = 'status-' + task.Status.toLowerCase();
                const statusIcon = getStatusIcon(task.Status);
                const pulseClass = task.Status === 'Running' ? 'pulse' : '';
                
                html += '<div class="task ' + statusClass + '">';
                html += '<div class="task-header">';
                html += '<div class="task-name">';
                html += '<span class="status-icon ' + pulseClass + '">' + statusIcon + '</span>';
                html += '<span>' + task.Description + '</span>';
                html += '</div>';
                html += '<div><span class="task-status">' + task.Status + ' (' + task.Runtime + 's)</span></div>';
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
            
            html += '</div></div>';
            html += '<div class="last-update">⏱️ Last updated: ' + data.LastUpdate + '</div>';
            
            document.getElementById('content').innerHTML = html;
        }
        
        updateDashboard(dashboardData);
    </script>
</body>
</html>
"@

    # Write with UTF-8 encoding to ensure emojis display correctly
    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($htmlPath, $html, $utf8NoBom)

    # Only launch browser if dashboard was just created AND LaunchBrowser was requested
    if ($LaunchBrowser) {
        $fullPath = Resolve-Path $htmlPath
        $null = Start-Process $fullPath
        Write-Host "Dashboard launched in browser..." -ForegroundColor Cyan
    }
    
    if (-not $Quiet) {
        Write-Host "Dashboard exported to: $htmlPath" -ForegroundColor Green
    }
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
