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
            $Script:PrimaryExchangePsSession = Get-PSSession |
                Where-Object { $_.Availability -eq "Available" -and
                    $_.ConfigurationName -eq "Microsoft.Exchange" -and
                    $_.State -eq "Opened" } |
                Select-Object -First 1
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
            Import-PSSession $Script:CachedExchangePsSession[$ServerName] -AllowClobber | Out-Null
            $params = @{
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
    param(
        [scriptblock]
        $CatchActionFunction
    )
    process {
        if ($null -ne $Script:CachedExchangePsSession) {
            $Script:CachedExchangePsSession.Keys |
                ForEach-Object {
                    Remove-PSSession $Script:CachedExchangePsSession[$_]
                }

            $Script:CachedExchangePsSession = $null

            if ($null -ne $Script:PrimaryExchangePsSession) {
                try {
                    Import-PSSession $Script:PrimaryExchangePsSession -ErrorAction Stop -AllowClobber | Out-Null
                } catch {
                    Invoke-CatchActionError $CatchActionFunction
                }
            }
        }
    }
}
