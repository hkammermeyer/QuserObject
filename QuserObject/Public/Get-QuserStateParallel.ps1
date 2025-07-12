#Requires -Module PoshRSJob
# The QuserObject module is assumed to be part of the same module manifest or available in the environment.

function Get-QuserStateParallel {
    <#
    .SYNOPSIS
        Queries user session state on remote computers in parallel.

    .DESCRIPTION
        Get-QuserStateParallel uses Get-Quser and the PoshRSJob module to query multiple computers simultaneously.
        It determines an overall state for each computer (e.g., Active, Idle, Empty) based on its collective user sessions.
        It also calculates the duration a server has been idle or the time remaining until it becomes idle.

        This function requires the PoshRSJob module to be installed.

    .PARAMETER ComputerName
        One or more computer names to query. Defaults to the local machine. Accepts pipeline input.

    .PARAMETER IdleTime
        The time threshold to consider a user session as idle. Defaults to 1 hour.

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

        Queries two computers and returns their state, including any relevant idle timer properties.

    .OUTPUT
        [PSCustomObject]
        An object for each computer with its calculated state and session details.
    #>
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    param(
        [parameter(
            ValueFromPipeline,
            ValueFromPipelineByPropertyName,
            Position = 0)]
        [alias('__ServerName', 'ServerName', 'Server', 'Computer', 'Name', 'CN')]
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
        if ($Clipboard.IsPresent) {
            $ComputerName = Get-Clipboard | Where-Object { $_ -and $_.Trim() -ne '' }
        }

        if (-not $KeepDuplicates.IsPresent) {
            # Standardize to lowercase and remove duplicates.
            $ComputerName = $ComputerName | ForEach-Object { $_.ToLower() } | Sort-Object | Get-Unique
        }

        if (-not $ComputerName) {
            Write-Error "No computer names were provided. Please specify names via the -ComputerName parameter, -Clipboard switch, or pipeline."
            return
        }

        # Use a GUID to ensure job names are unique for each run.
        $jobGuid = (New-Guid).ToString()
    }

    process {
        # --- Parallel Job Execution ---
        $ComputerName | Start-RSJob -Name "QuserStateJob-$jobGuid" -Throttle $Throttle -ScriptBlock {
            param($ComputerName_Item)

            # Clear previous errors on this thread.
            $Error.Clear()

            $sessions = Get-Quser -Server $ComputerName_Item -WarningAction SilentlyContinue
            $timestamp = Get-Date

            # --- Result Analysis ---
            if ($sessions) {
                # --- State Logic ---

                # Define a threshold for an unreasonably large idle time to detect the quser.exe bug.
                $buggedIdleTimeThreshold = [timespan]::FromDays(365)

                # Condition 1: Check for any active users. A user is considered active if their state is 'Active' and their idle time
                # is either below the threshold OR above the bugged threshold (working around the quser bug).
                $hasActiveUsers = $sessions | Where-Object {
                    $_.State -eq 'Active' -and (
                        ($null -eq $_.IdleTime) -or
                        ($_.IdleTime -lt [timespan]$using:IdleTime) -or
                        ($_.IdleTime -gt $buggedIdleTimeThreshold)
                    )
                }

                $areAllSessionsIdle = -not ($sessions | Where-Object { $_.IdleTime -lt [timespan]$using:IdleTime })
                $areAllDisconnected = -not ($sessions | Where-Object { $_.State -ne 'Disc' })

                # --- Build Output Object ---
                $resultObject = [PSCustomObject]@{
                    Name         = $ComputerName_Item
                    State        = 'Mixed' # Default state
                    IdleDuration = $null
                    TimeToIdle   = $null
                    Timestamp    = $timestamp
                    SessionCount = @($sessions).Count
                    Session      = $sessions
                }

                if ($hasActiveUsers) {
                    $resultObject.State = 'Active'
                }
                elseif ($areAllSessionsIdle) {
                    $resultObject.State = 'Idle'
                    # Calculate the server's total idle duration.
                    # This is based on the session that has been idle the least amount of time.
                    $minIdleTime = ($sessions.IdleTime | Measure-Object -Minimum).Minimum
                    $resultObject.IdleDuration = $minIdleTime
                }
                elseif ($areAllDisconnected) {
                    $resultObject.State = 'Disc'
                }

                # For 'Disc' or 'Mixed' states, calculate the time remaining until the server becomes idle.
                if ($resultObject.State -in @('Disc', 'Mixed')) {
                    $minIdleTime = ($sessions.IdleTime | Where-Object { $_ } | Measure-Object -Minimum).Minimum
                    if ($null -ne $minIdleTime) {
                         $resultObject.TimeToIdle = $using:IdleTime - $minIdleTime
                    }
                }

                return $resultObject
            }
            else {
                # This block handles cases where Get-Quser returned no sessions, either due to an error or an empty server.
                $lastError = $Error[0]
                $state = if ($lastError -and $lastError.Exception.Message -like '*No User exists for*') {
                    'Empty'
                }
                elseif ($lastError -and $lastError.Exception.Message -like '*The RPC server is unavailable*') {
                    'RPC Unavailable'
                }
                else {
                    "Error: $($lastError.Exception.GetBaseException().Message)"
                }

                # Build the error/empty object
                [PSCustomObject]@{
                    Name         = $ComputerName_Item
                    State        = $state
                    IdleDuration = $null
                    TimeToIdle   = $null
                    Timestamp    = $timestamp
                    SessionCount = 0
                    Session      = $null
                }
            }
        } | Out-Null # Suppress job objects from Start-RSJob output
    }

    end {
        # --- Job Cleanup and Output ---
        $allJobs = Get-RSJob -Name "QuserStateJob-$jobGuid"
        if ($allJobs) {
            $results = $allJobs | Wait-RSJob -ShowProgress -Timeout $QueryTimeout | Receive-RSJob
            $allJobs | Remove-RSJob -Force

            # Return final results, sorted by name for consistency.
            return $results | Sort-Object -Property Name
        }
    }
}
