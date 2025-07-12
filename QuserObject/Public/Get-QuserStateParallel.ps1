#Requires -Module QuserObject, PoshRSJob

function Get-QuserStateParallel {
    [CmdletBinding()]
    param(
        [parameter(
            ValueFromPipeline,
            ValueFromPipelineByPropertyName,
            Position = 0)]
        [alias('Name')]
        [string[]]$ComputerName,
        [timespan]$IdleTime = '01:00:00',
        [switch]$Clipboard,
        [switch]$KeepDuplicates,
        [int]$Throttle = 64,
        [int]$QueryTimeout = 600
    )
    begin {
        # Process clipboard items if clipboard switch is used.
        if ($Clipboard.IsPresent) {
            # Remove empty entries
            $ComputerName = Get-Clipboard | Where-Object { $PSItem }
        }
        # Remove duplicates
        if (-not $KeepDuplicates.IsPresent) {
            $ComputerName = $ComputerName.ToLower() | Select-Object -Unique
        }
    }
    process {
        $ComputerName | Start-RSJob -Name 'QuserStateJob' -ModulesToImport "$PSScriptRoot\lib\QuserObject" -Throttle $Throttle -ScriptBlock {
            param($ComputerName_Item)
            Clear-Variable -Name Result, Timestamp -ErrorAction Ignore

            $Result = Get-Quser -Server $ComputerName_Item -WarningAction SilentlyContinue |
                Select-Object -Property Id, UserName, LogonTime, State, IdleTime, SessionName, Server
            $Timestamp = Get-Date

            if ($Result) {
                # All users disconnected, $IdleTime ignored
                if ($Result.State -notcontains 'Active') {
                    [PSCustomObject]@{
                        Name         = $ComputerName_Item
                        State        = 'Disc'
                        Timestamp    = $Timestamp
                        SessionCount = @($Result).count
                        Session      = $Result
                    }
                }
                # Less then $IdleTime or at least one user active
                elseif (($Result.IdleTime -lt [timespan]$using:IdleTime) -or ($Result.IdleTime -contains $null)) {
                    [PSCustomObject]@{
                        Name         = $ComputerName_Item
                        State        = 'Active'
                        Timestamp    = $Timestamp
                        SessionCount = @($Result).count
                        Session      = $Result
                    }
                }
                # All users idle at determined $IdleTime
                elseif (-not ($Result.IdleTime -lt [timespan]$using:IdleTime)) {
                    [PSCustomObject]@{
                        Name         = $ComputerName_Item
                        State        = 'Idle'
                        Timestamp    = $Timestamp
                        SessionCount = @($Result).count
                        Session      = $Result
                    }
                }
                # Something is not right.
                else {
                    Write-Warning "[$ComputerName_Item] Something went wrong. Please report issue."
                    $Result | Sort-Object State, IdleTime | Format-Table -AutoSize
                }
                <#
                # Display quser results if debug is enabled
                if ($PSCmdlet.MyInvocation.BoundParameters['Debug'].IsPresent) {
                    $Result | Sort-Object State, IdleTime | Format-Table -AutoSize
                }#>
            }
            elseif (-not $Result) {
                # No users on server
                if ($Error[0].Exception.Message -like "*No User exists for *") {
                    [PSCustomObject]@{
                        Name         = $ComputerName_Item
                        State        = 'Empty'
                        Timestamp    = $Timestamp
                        SessionCount = '0'
                        Session      = ''
                    }
                }
                # Assume RPC server is unavailable
                else {
                    [PSCustomObject]@{
                        Name         = $ComputerName_Item
                        State        = 'The RPC server is unavailable.'
                        Timestamp    = $Timestamp
                        SessionCount = ''
                        Session      = ''
                    }
                }
            } #else if
        } | Out-Null #Start-RSJob ScriptBlock
    } #process
    end {
        Get-RSJob -Name 'QuserStateJob' | Wait-RSJob -ShowProgress -Timeout $QueryTimeout | Receive-RSJob
        Get-RSJob -Name 'QuserStateJob' | Remove-RSJob -Force
    }
} #function