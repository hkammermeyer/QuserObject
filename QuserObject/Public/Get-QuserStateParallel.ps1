#Requires -Module PoshRSJob
# The QuserObject module is assumed to be part of the same module manifest or available in the environment.

function Get-QuserStateParallel {
    <#
    .SYNOPSIS
        Queries user session state on remote computers in parallel.

    .DESCRIPTION
        Get-QuserStateParallel uses Get-Quser and the PoshRSJob module to query multiple computers simultaneously.
        It determines an overall state for each computer (e.g., Active, Idle, Empty) based on its collective user sessions.

        This function requires the PoshRSJob module. You can install it by running:
        Install-Module -Name PoshRSJob -Scope CurrentUser

    .PARAMETER ComputerName
        One or more computer names to query. Accepts pipeline input.

    .PARAMETER IdleTime
        The time threshold to consider a user session as idle. Defaults to 1 hour. Any session (including disconnected) with an idle time
        at or above this value will contribute to an overall 'Idle' state.

    .PARAMETER Clipboard
        A switch to get computer names from the system clipboard.

    .PARAMETER KeepDuplicates
        A switch to prevent the removal of duplicate computer names from the input list.

    .PARAMETER Throttle
        The maximum number of parallel jobs to run at once. Defaults to 64.

    .PARAMETER QueryTimeout
        The maximum time in seconds to wait for all jobs to complete. Defaults to 600 (10 minutes).

    .EXAMPLE
        PS C:\> Get-QuserStateParallel -ComputerName 'SERVER01', 'WKSTN05'

        Name            State           Timestamp             SessionCount Session
        ----            -----           ---------             ------------ -------
        SERVER01        Active          7/12/2025 11:09:15 AM            2 {@{Id=1; UserName=jdoe; ...}, @{Id=2; UserName=asmith; ...}}
        WKSTN05         Empty           7/12/2025 11:09:16 AM            0

        Queries two computers and returns their state.

    .EXAMPLE
        PS C:\> 'SERVER01', 'SERVER02', 'SERVER03' | Get-QuserStateParallel -IdleTime '00:15:00' -Throttle 32

        Pipes computer names to the function, setting a custom idle threshold of 15 minutes and a throttle limit of 32.

    .EXAMPLE
        PS C:\> Get-QuserStateParallel -Clipboard

        Runs a query against the list of computer names currently on the clipboard.

    .OUTPUT
        [PSCustomObject]
        An object for each computer with its Name, overall State, Timestamp of the query, SessionCount, and the detailed Session object(s).
    #>
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    param(
        [parameter(
            ValueFromPipeline,
            ValueFromPipelineByPropertyName,
            Position = 0)]
        [alias('Name')]
        [ValidateNotNullOrEmpty()]
        [string[]]$ComputerName = $env:computername,

        [timespan]$IdleTime = '01:00:00',
        [switch]$Clipboard,
        [switch]$KeepDuplicates,
        [int]$Throttle = 64,
        [int]$QueryTimeout = 600
    )

    begin {
        # --- Input Processing ---
        # If the -Clipboard switch is used, get computer names from the clipboard.
        # Filter out any empty or whitespace-only lines.
        if ($Clipboard.IsPresent) {
            $ComputerName = Get-Clipboard | Where-Object { $_ -and $_.Trim() -ne '' }
        }

        # Unless specified otherwise, convert all names to lowercase and remove duplicates for efficiency.
        if (-not $KeepDuplicates.IsPresent) {
            $ComputerName = $ComputerName | ForEach-Object { $_.ToLower() } | Sort-Object | Get-Unique
        }

        # If after processing input there are no computer names, stop execution.
        if (-not $ComputerName) {
            Write-Error "No computer names were provided. Please specify names via the -ComputerName parameter, -Clipboard switch, or pipeline."
            return
        }

        # Generate a unique ID for this run to prevent job name collisions if the function
        # is called multiple times in the same session.
        $jobGuid = (New-Guid).ToString()
    }

    process {
        # --- Parallel Job Execution ---
        # Pipe each computer name to Start-RSJob to process them in parallel.
        $ComputerName | Start-RSJob -Name "QuserStateJob-$jobGuid" -Throttle $Throttle -ScriptBlock {
            # This scriptblock runs on a separate thread for each computer.
            param($ComputerName_Item)

            # Clear the thread's error buffer to ensure we only check for errors from the Get-Quser command.
            $Error.Clear()

            # Execute the query. We use -WarningAction SilentlyContinue to suppress non-terminating warnings
            # from quser.exe that would otherwise clutter the console.
            $sessions = Get-Quser -Server $ComputerName_Item -WarningAction SilentlyContinue

            # Capture the timestamp immediately after the query.
            $timestamp = Get-Date

            # --- Result Analysis ---
            if ($sessions) {
                # This block runs if Get-Quser returned one or more session objects.

                # --- State Logic ---
                # The goal is to determine a single, overall state for the computer based on its sessions.

                # Condition 1: Is there at least one 'Active' user whose idle time is below the threshold?
                # This is the primary indicator of an active machine.
                $hasActiveUsers = $sessions | Where-Object { $_.State -eq 'Active' -and ($null -eq $_.IdleTime -or $_.IdleTime -lt [timespan]$using:IdleTime) }

                # Condition 2: Are ALL sessions (including 'Disc') idle for longer than the threshold?
                # This defines the 'Idle' state. A session is considered idle if its IdleTime exceeds the threshold, regardless of its state.
                $areAllSessionsIdle = -not ($sessions | Where-Object { $_.IdleTime -lt [timespan]$using:IdleTime })

                # Condition 3: Are ALL sessions in the 'Disc' state?
                # This is a fallback for when a machine has only disconnected users who have NOT yet met the idle threshold.
                $areAllDisconnected = -not ($sessions | Where-Object { $_.State -ne 'Disc' })

                # Build the output object.
                $resultObject = [PSCustomObject]@{
                    Name         = $ComputerName_Item
                    State        = 'Mixed' # Default to 'Mixed' and overwrite based on the logic below.
                    Timestamp    = $timestamp
                    SessionCount = @($sessions).Count
                    Session      = $sessions
                }

                if ($hasActiveUsers) {
                    $resultObject.State = 'Active'
                }
                elseif ($areAllSessionsIdle) {
                    $resultObject.State = 'Idle'
                }
                elseif ($areAllDisconnected) {
                    $resultObject.State = 'Disc'
                }
                # If none of the above specific states match, the state remains 'Mixed'.

                return $resultObject
            }
            else {
                # This block runs if Get-Quser returned nothing ($null), which indicates either an error or no sessions.
                # We inspect the $Error automatic variable because native executable errors (like from quser.exe)
                # often don't throw terminating script errors that a try/catch block would handle.

                $lastError = $Error[0]

                if ($lastError -and $lastError.Exception.Message -like '*No User exists for*') {
                    # The command succeeded but found no logged-on users.
                    [PSCustomObject]@{
                        Name         = $ComputerName_Item
                        State        = 'Empty'
                        Timestamp    = $timestamp
                        SessionCount = 0
                        Session      = $null
                    }
                }
                elseif ($lastError -and $lastError.Exception.Message -like '*The RPC server is unavailable*') {
                    # A common network/firewall error indicating the remote machine could not be contacted.
                    [PSCustomObject]@{
                        Name         = $ComputerName_Item
                        State        = 'RPC Unavailable'
                        Timestamp    = $timestamp
                        SessionCount = 0
                        Session      = $null
                    }
                }
                else {
                    # A different, unexpected error occurred.
                    [PSCustomObject]@{
                        Name         = $ComputerName_Item
                        State        = "Error: $($lastError.Exception.GetBaseException().Message)"
                        Timestamp    = $timestamp
                        SessionCount = 0
                        Session      = $null
                    }
                }
            }
        } | Out-Null # Suppress the job objects that Start-RSJob outputs to the pipeline.
    }

    end {
        # --- Job Cleanup and Output ---
        # Wait for all jobs created during this run to complete.
        $allJobs = Get-RSJob -Name "QuserStateJob-$jobGuid"
        if ($allJobs) {
            # Wait for jobs to finish, showing a progress bar, and then receive the results.
            $results = $allJobs | Wait-RSJob -ShowProgress -Timeout $QueryTimeout | Receive-RSJob

            # Clean up by removing the completed jobs.
            $allJobs | Remove-RSJob -Force

            # Return the collected results, sorted by name for consistent output.
            return $results | Sort-Object -Property Name
        }
    }
}
