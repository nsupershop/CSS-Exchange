# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

function Get-ApplicationHostConfig {
    [CmdletBinding()]
    [OutputType([System.Xml.XmlNode])]
    param ()

    $appHostConfig = New-Object -TypeName Xml
    try {
        $appHostConfigPath = "$($env:WINDIR)\System32\inetsrv\config\applicationHost.config"
        $appHostConfig.Load($appHostConfigPath)
    } catch {
        Write-Verbose "Failed to loaded 'applicationHost.config' file. $_"
        $appHostConfig = $null
    }
    return $appHostConfig
}
