# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

function Get-IISModules {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNode]$ApplicationHostConfig,

        [Parameter(Mandatory = $false)]
        [scriptblock]$CatchActionFunction
    )

    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        function ConvertCmdEnvironmentVariableToFullPath {
            [CmdletBinding()]
            [OutputType([System.String])]
            param(
                [String]$PathWithVar
            )

            $returnPath = $null
            if (-not([String]::IsNullOrEmpty($PathWithVar))) {
                # Assuming that we have the env var always at the beginning of the string and no other vars within the string
                # Example: %windir%\system32\someexample.dll
                $preparedPath = ($PathWithVar.Split("%", [System.StringSplitOptions]::RemoveEmptyEntries))
                if ($preparedPath.Count -eq 2) {
                    if ($preparedPath[0] -notmatch "\\.+\\") {
                        $varPath = [System.Environment]::GetEnvironmentVariable($preparedPath[0])
                        $returnPath = [String]::Join("", $varPath, $($preparedPath[1]))
                    }
                }
            }
            return $returnPath
        }

        function GetIISModulesLoaded {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory = $true)]
                [System.Xml.XmlNode]$Xml
            )
            process {
                try {
                    $iisModulesList = New-Object 'System.Collections.Generic.List[object]'
                    $globalModules = $Xml.configuration.'system.webServer'.globalModules.add
                    if ($globalModules.Count -ge 1) {
                        Write-Verbose "At least one module is loaded by IIS"
                        foreach ($m in $globalModules) {
                            Write-Verbose "Now processing module: $($m.Name)"
                            $mInfoObject = [PSCustomObject]@{
                                Name             = $m.Name
                                Path             = $null
                                Signed           = $false
                                SignatureDetails = $null
                            }

                            $moduleFilePath = $m.Image
                            if ($moduleFilePath -match "\%.+\%") {
                                # Overwrite path when environment variables were found
                                Write-Verbose "Environment variable found in path: $moduleFilePath"
                                $moduleFilePath = ConvertCmdEnvironmentVariableToFullPath -PathWithVar $m.Image
                            }

                            $mInfoObject.Path = $moduleFilePath

                            try {
                                Write-Verbose "Querying file signing information"
                                $signature = Get-AuthenticodeSignature -FilePath $moduleFilePath -ErrorAction Stop
                                # Signature Status Enum Values:
                                # <0> Valid, <1> UnknownError, <2> NotSigned, <3> HashMismatch,
                                # <4> NotTrusted, <5> NotSupportedFileFormat, <6> Incompatible,
                                # https://docs.microsoft.com/dotnet/api/system.management.automation.signaturestatus
                                if (($signature.Status -ne 1) -and
                                    ($signature.Status -ne 2) -and
                                    ($signature.Status -ne 5) -and
                                    ($signature.Status -ne 6)) {
                                    Write-Verbose "Signature information found. Status: $($signature.Status)"

                                    $mInfoObject.Signed = $true
                                    $mDetails = [PSCustomObject]@{
                                        Signer            = $null
                                        SignatureStatus   = $signature.Status
                                        IsMicrosoftSigned = $false
                                    }

                                    if ($null -ne $signature.SignerCertificate.Subject) {
                                        Write-Verbose "Signer information found. Subject: $($signature.SignerCertificate.Subject)"
                                        $mDetails.Signer = $signature.SignerCertificate.Subject.ToString()
                                        $mDetails.IsMicrosoftSigned = $signature.SignerCertificate.Subject -cmatch "O=Microsoft Corporation, L=Redmond, S=Washington"
                                    }

                                    $mInfoObject.SignatureDetails = $mDetails
                                }

                                $iisModulesList.Add($mInfoObject)
                            } catch {
                                Write-Verbose "Unable to validate file signing information"
                                if ($null -ne $CatchActionFunction) {
                                    & $CatchActionFunction
                                }
                            }
                        }
                    } else {
                        Write-Verbose "No modules are loaded by IIS"
                    }
                } catch {
                    Write-Verbose "Failed to process global module information. $_"
                    if ($null -ne $CatchActionFunction) {
                        & $CatchActionFunction
                    }
                }
            }
            end {
                return $iisModulesList
            }
        }
    }
    process {
        $modules = GetIISModulesLoaded -Xml $ApplicationHostConfig

        # Validate that all modules are signed by Microsoft Corp.
        if ($modules.SignatureDetails.IsMicrosoftSigned.Contains($false)) {
            Write-Verbose "At least one module which is loaded by IIS is not digitally signed by 'Microsoft Corporation'"
            $allModulesSignedByMSFT = $false
        } else {
            Write-Verbose "All modules which are loaded by IIS are signed by 'Microsoft Corporation'"
            $allModulesSignedByMSFT = $true
        }

        # Validate if all signatures are valid (regardless of whether signed by Microsoft Corp. or not)
        $allSignaturesValid = $true
        foreach ($m in $modules) {
            Write-Verbose "Module: $($m.Name) Signature Status: $($m.SignatureDetails.SignatureStatus)"
            if (($m.Signed -eq $true) -and
                ($m.SignatureDetails.SignatureStatus -ne 0)) {
                $allSignaturesValid = $false
            }
        }

        # Validate if all modules that are loaded are digitally signed
        if ($modules.Signed.Contains($false)) {
            Write-Verbose "At least one module which is loaded by IIS is not digitally signed"
            $allModulesAreSigned = $false
        } else {
            Write-Verbose "All modules which are loaded by IIS are digitally signed"
            $allModulesAreSigned = $true
        }
    }
    end {
        return [PSCustomObject]@{
            AllSignedModulesSignedByMSFT = $allModulesSignedByMSFT
            AllSignaturesValid           = $allSignaturesValid
            AllModulesSigned             = $allModulesAreSigned
            ModuleList                   = $modules
        }
    }
}
