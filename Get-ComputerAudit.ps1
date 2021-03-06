Function Get-ComputerAudit
{

    [CmdletBinding()]
    Param
    (

       [Parameter(Position=0)]
       [string[]]$ComputerName = 'localhost'

    )

    BEGIN
    {

        Class ComputerAudit
        {

            [string]$ComputerName
            [string]$IPv4Address = ''
            [string]$MacAddress = ''
            [string]$LastBootUpTime = ''
            [string]$Manufacturer = ''
            [string]$Model = ''
            [string]$SerialNumber = ''
            [string]$OperatingSystem = ''
            [string]$OperatingSystemVersion = ''
            [string]$OSServicePack = ''
            [string]$OSArchitecture = ''
            [sbyte]$OSLicenseStatus = -1          
            [string]$PSVersion = '-1'
            [sbyte]$SMB1Enabled = -1
            
        }

    }

    PROCESS
    {

        ForEach ( $Computer in $ComputerName )
        { 

            ###################################
            # Step 1 - Create computer object #
            ###################################

            $ComputerAudit = New-Object -TypeName ComputerAudit
            $ComputerAudit.ComputerName = $Computer
            
            $TestConnect = Test-Connection $Computer -Count 1 -ErrorAction SilentlyContinue

            if ( $TestConnect )
            {

                $ComputerAudit.IPv4Address = $TestConnect.IPV4Address.IPAddressToString

            }
            else
            {

                Write-Error -Message "Unable to connect to target '$Computer'." -Exception ([System.TimeoutException]::new()) -Category ConnectionError -CategoryTargetName $Computer
                Continue
                
            }
            
            #####################################
            # Step 2 - Get Information from WMI #
            #####################################

            # Get WMI Properties
            # At this stage always use Select-Object so that as little is held in memory as possible
            
            Write-Verbose "Querying WMI classes on target '$Computer'."
            
            Try
            {

                ## Win32_ComputerSystem
                $WmiComputerSystem = Get-WmiObject Win32_ComputerSystem -ComputerName $Computer -ErrorAction Stop | Select-Object Manufacturer,Model

                $ComputerAudit.Manufacturer = $WmiComputerSystem.Manufacturer
                $ComputerAudit.Model = $WmiComputerSystem.Model


                ## Win32_Bios
                $WmiBios = Get-WmiObject Win32_Bios -ComputerName $Computer -ErrorAction Stop | Select-Object SerialNumber

                $ComputerAudit.SerialNumber = $WmiBios.SerialNumber


                ## Win32_OperatingSystem
                $WmiOsDetails = Get-WmiObject Win32_OperatingSystem -ComputerName $Computer -ErrorAction Stop |
                    Select-Object Caption,Version,ServicePackMajorVersion,OSArchitecture,@{label='LastBootUpTime';expression={$_.ConverttoDateTime($_.lastbootuptime)}}

                $ComputerAudit.OperatingSystem = $WmiOSDetails.Caption
                $ComputerAudit.OperatingSystemVersion = $WmiOSDetails.Version
                $ComputerAudit.OSServicePack = $WmiOSDetails.ServicePackMajorVersion
                $ComputerAudit.OSArchitecture = $WmiOSDetails.OSArchitecture
                $ComputerAudit.LastBootUpTime =  $WmiOSDetails.LastBootUpTime.ToString()


                ## SoftwareLicensingProduct
                ## Warning, hardcoded ApplicationId, in tests this always returned OS licence but test was limited to specific environment but run against server OS and client OS
                $WmiFilter = "ApplicationId='55c92734-d682-4d71-983e-d6ec3f16059f' AND PartialProductKey IS NOT NULL"
                $WMISoftwareLicensing = Get-WmiObject SoftwareLicensingProduct -ComputerName $Computer -Filter $WmiFilter -ErrorAction Stop |
                    Select-Object LicenseStatus

                $ComputerAudit.OSLicenseStatus = $WMISoftwareLicensing.LicenseStatus


                # Win32_NetworkAdapterConfiguration (MAC Address)
                # Excludes Servicename Netft, which is 'Failover Cluster Virtual Adapter' 
                $MacAddress = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $Computer -ErrorAction Stop |
                    Where-Object { ($_.IPEnabled -eq $true) -and ($_.ServiceName -ne 'Netft') } | Select-Object -ExpandProperty MacAddress
        
                $ComputerAudit.MacAddress = $MacAddress -join ','

            }
            Catch
            {

                Write-Error $_.Exception.Message -ErrorId $_.FullyQualifiedErrorId

            }


            ##############################################
            # Step 3 - Get information from the Registry #
            ##############################################

            $RemoteRegError = $null
            
            # Check RemoteRegistry Service
            Try
            {
            
                Write-Verbose "Connecting to RemoteRegistry service on target '$Computer'."
                $RRService = Get-Service RemoteRegistry -ComputerName $Computer -ErrorAction Stop
            
            }
            Catch
            {

                $RemoteRegError = $_
                                
            }



            If ( ($RRService.Status -eq 'Running') -and ($RemoteRegError -eq $null) )
            {
                
                # Open the registry
                try
                {

                    $Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Computer)

                }
                catch
                {

                    Write-Error $_.Exception.Message -ErrorId $_.FullyQualifiedErrorId

                }
            

                # Get details from Registry
                try
                {

                    Write-Verbose "Querying RemoteRegistry on target '$Computer'."

                    # Get Powershell Version
                    $RegKeyPSVersion = $Reg.OpenSubKey("SOFTWARE\Microsoft\PowerShell\3\PowerShellEngine")
            
                    If ( $RegKeyPSVersion -eq $null )
                    {
            
                        $RegKeyPSVersion = $Reg.OpenSubKey("SOFTWARE\Microsoft\PowerShell\1\PowerShellEngine")    
            
                    }
            
                    $ComputerAudit.PSVersion = $RegKeyPSVersion.GetValue("PowerShellVersion")
                    $RegKeyPSVersion.Close()
                
                    # Get SMB 1
                    $RegKeySmb1 = $Reg.OpenSubKey("SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters")
                    
                    $Smb1Value = $RegKeySmb1.GetValue("SMB1")
                    $RegKeySmb1.Close()
            
                    If ( $Smb1Value -ne $null )
                    {
            
                        $ComputerAudit.SMB1Enabled = $Smb1Value
            
                    }
                    Else
                    {
            
                        $ComputerAudit.SMB1Enabled = 1
            
                    }

                }#EndOfTry

                catch
                {
            
                    Write-Error $_.Exception.Message -ErrorId $_.FullyQualifiedErrorId
            
                }
                
                finally
                {
                
                    $Reg.Close()
                    $Reg.Dispose()
                
                }

            } #End the statement "if remote registry service is working"
            ElseIf ($RRService.Status -ne 'Running')
            {
             
                Write-Warning "Registry service on '$Computer' is not running, its status is '$($RRService.Status)'."

            }
            ElseIf ($RemoteRegError.Status -ne $null)
            {
             
                Write-Error $RemoteRegError.Exception.Message -ErrorId $RemoteRegError.FullyQualifiedErrorId

            }


            #################################
            # Output a ComputerAudit object #
            #################################
        
            Write-Output $ComputerAudit
    
            Clear-Variable ComputerAudit
         
        }#EndOfProcessForEach

    }

    END
    {

        #No action in end block

    }

}#EndOfFunction