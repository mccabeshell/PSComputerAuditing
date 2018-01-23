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
            
            If ( -not (Test-Connection $Computer -Quiet -Count 1) )
            {

                Write-Error "Cannot connect to target '$Computer'."
                Continue
                
            }
            
            #####################################
            # Step 2 - Get Information from WMI #
            #####################################

            # Get WMI Properties
            ## At this stage always use Select-Object so that as little is held in memory as possible
            
            Write-Verbose "Querying WMI classes on target '$ComputerName'."
            
            Try
            {

                ## Win32_ComputerSystem
                $WmiComputerSystem = Get-WmiObject Win32_ComputerSystem -ErrorAction Stop | Select-Object Manufacturer,Model

                ## Win32_Bios
                $WmiBios = Get-WmiObject Win32_Bios -ComputerName $Computer -ErrorAction Stop | Select-Object SerialNumber

                ## Win32_OperatingSystem
                $WmiOsDetails = Get-WmiObject Win32_OperatingSystem -ComputerName $Computer -ErrorAction Stop |
                    Select-Object Caption,Version,ServicePackMajorVersion,OSArchitecture,@{label='LastBootUpTime';expression={$_.ConverttoDateTime($_.lastbootuptime)}}

                ## SoftwareLicensingProduct
                ### Warning, hardcoded ApplicationId, in tests this always returned OS licence but test was limited to specific environment but run against server OS and client OS
                $WMISoftwareLicensing = Get-WmiObject SoftwareLicensingProduct -ComputerName $Computer -Filter "ApplicationId='55c92734-d682-4d71-983e-d6ec3f16059f' AND PartialProductKey IS NOT NULL" -ErrorAction Stop |
                    Select-Object LicenseStatus

            }
            Catch
            {

                $ComputerAudit.LicenseStatus = -1

            }

            # Populate Computer Audit Properties with WMI details
            try
            {

                $ComputerAudit.Manufacturer = $WmiComputerSystem.Manufacturer
                $ComputerAudit.Model = $WmiComputerSystem.Model
                $ComputerAudit.SerialNumber = $WmiBios.SerialNumber
                $ComputerAudit.OperatingSystem = $WmiOSDetails.Caption
                $ComputerAudit.OperatingSystemVersion = $WmiOSDetails.Version
                $ComputerAudit.OSServicePack = $WmiOSDetails.ServicePackMajorVersion
                $ComputerAudit.OSArchitecture = $WmiOSDetails.OSArchitecture
                $ComputerAudit.LastBootUpTime =  $WmiOSDetails.LastBootUpTime
                $ComputerAudit.OSLicenseStatus = $WMISoftwareLicensing.LicenseStatus

            }
            catch
            {
             
                throw $_

            }


            ##############################################
            # Step 3 - Get information from the Registry #
            ##############################################

            $RemoteRegError = 0
            
            # Check RemoteRegistry Service
            Try
            {
            
                Write-Verbose "Checking status of RemoteRegistry service on target '$ComputerName'."
                $RRService = Get-Service RemoteRegistry -ComputerName $Computer -ErrorAction Stop
            
            }
            Catch
            {

                $RemoteRegError = 1
                                
            }



            If ( ($RRService.Status -eq 'Running') -and ($RemoteRegError -eq 0) )
            {
                
                # Open the registry
                try
                {

                    $Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Computer)
                    Write-Verbose "Connecting to RemoteRegistry on target '$ComputerName'."

                }
                catch
                {

                    Throw $_

                }
            

                # Get details from Registry
                try
                {

                    Write-Verbose "Querying RemoteRegistry on target '$ComputerName'."

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
            
                    Throw $_
            
                }
                
                finally
                {
                
                    $Reg.Close()
                    $Reg.Dispose()
                
                }

            } #End the statement "if remote registry service is working"
            Else
            {
             
                Write-Warning "Unable to verify if remote registry service is running on '$Computer'. Results do not include information from the registry"

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