<#
    .Synopsis

        Get collective terminal session status.

    .Parameter Server

        The server to be queried. Default is current.

    .Parameter IdleTime

        By default, IdleTime is set to '01:00:00' (1 hour).

    .Parameter KeepDuplicateServer

        By default duplicate servers are removed, using switch will keep all servers from input.

    .Parameter Clipboard

        Clipboard switch will process items from clipboard.

    .Example

        Get-QuserState

    .Example

        Get-QuserState -Server 'ThisServer'

    .Example

        Get-QuserState -Server 'ThisServer', 'ThatServer'

    .Example

        Get-ADComputer 'ThisServer' | Get-QuserState

    .NOTES

        Author: Harold Kammermeyer
#>
function Get-QuserState {
    [CmdletBinding()]
    [Alias('Get-ServerState')]
    param(
        [parameter(
            ValueFromPipeline,
            ValueFromPipelineByPropertyName,
            Position = 0)]
        [alias('__ServerName', 'ServerName', 'Computer', 'Name', 'ComputerName', 'CN')]
        [string[]]
        $Server = $env:computername,
        [timespan]
        $IdleTime = '01:00:00',
        [switch]
        $KeepDuplicateServer,
        [switch]
        $Clipboard
    )
    begin {
        # Process clipboard items if clipboard switch is used
        if ($Clipboard.IsPresent) {
            $Server = Get-Clipboard | Where-Object { $PSItem }
        }
        # Remove duplicates
        if (-not $KeepDuplicateServer.IsPresent) {
            $Server = $Server.ToLower() | Select-Object -Unique
        }
    }

    process {
        foreach ($Server_Item in $Server) {
            if ($Result) { Clear-Variable -Name Result }
            $Result = Get-Quser -Server $Server_Item -WarningAction SilentlyContinue

            if ($Result) {
                # All sessions in disconnected state
                if ($Result.State -notcontains 'Active') {
                    [PSCustomObject]@{
                        Name  = $Server_Item
                        State = 'Disc'
                    }
                }
                # All sessions in idle ($IdleTime) state
                elseif (($Result.IdleTime -ge $IdleTime) -and ($Result.Sessionname -notcontains 'console')) {
                    [PSCustomObject]@{
                        Name  = $Server_Item
                        State = 'Idle'
                    }
                }
                # All sessions in active state
                else {
                    [PSCustomObject]@{
                        Name  = $Server_Item
                        State = 'Active'
                    }
                }
            } #if
            elseif (-not $Result) {
                # No users on server
                if ($Error[0].Exception.Message -like '*No User exists for *') {
                    [PSCustomObject]@{
                        Name  = $Server_Item
                        State = 'Empty'
                    }
                }
                # Misc failures
                else {
                    [PSCustomObject]@{
                        Name  = $Server_Item
                        State = $Error[0].Exception.Message
                    }
                }
            } #elseif
        } #foreach
    } #process

    end {
        #intentionally left blank
    }
} #function
