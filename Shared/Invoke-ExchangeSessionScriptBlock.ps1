# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

. $PSScriptRoot\Invoke-CatchActionError.ps1
function Invoke-ExchangeSessionScriptBlock {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]
        $ServerName,

        [Parameter(Mandatory = $true)]
        [scriptblock]
        $ScriptBlock,

        [object]
        $ArgumentList,

        [scriptblock]
        $CatchActionFunction
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"

        if ($null -eq $Script:CachedExchangePsSession) {
            $Script:CachedExchangePsSession = @{}
        }

        if (-not ($Script:CachedExchangePsSession.ContainsKey($ServerName))) {
            try {
                $params = @{
                    ConfigurationName = "Microsoft.Exchange"
                    ConnectionUri     = "http://$ServerName/powershell"
                    Authentication    = "Kerberos"
                    ErrorAction       = "Stop"
                }
                $session = New-PSSession @params
                $Script:CachedExchangePsSession.Add($ServerName, $session)
            } catch {
                Invoke-CatchActionError $CatchActionFunction
            }
        }
    }
    process {
        try {
            $params = @{
                Session      = $Script:CachedExchangePsSession[$ServerName]
                ScriptBlock  = $ScriptBlock
                ArgumentList = $ArgumentList
                ErrorAction  = "Stop"
            }
            return Invoke-Command @params
        } catch {
            Invoke-CatchActionError $CatchActionFunction
        }
    }
}

function Invoke-RemoveCachedExchangeSessions {
    [CmdletBinding()]
    param()
    process {
        if ($null -ne $Script:CachedExchangePsSession) {
            $Script:CachedExchangePsSession.Keys |
                ForEach-Object {
                    Remove-PSSession $Script:CachedExchangePsSession[$_]
                }

            $Script:CachedExchangePsSession = $null
        }
    }
}
