Class ComputerAudit
{

    # Properties

    ## Properties Available in Multiple Sources
    [string]$ComputerName
    [string]$DNSName
    [string]$IPAddress
    [string]$OperatingSystem
    [string]$OperatingSystemVersion
    [string]$OperatingSystemServicePack
    [string]$Manufacturer
    [string]$Model
    
    ## Properties Generally from Remote Connection
    [string]$ConnectionSuccessful
    [string]$MacAddress
    [string]$LastBootUpTime
    [string]$SerialNumber
    [string]$OSArchitecture
    [sbyte]$OSLicenseStatus = -1          
    [string]$PSVersion
    [sbyte]$SMB1Enabled = -1

    ## Active Directory Properties
    [string]$ADEnabled
    [datetime]$ADLastLogonTime

    ## Wsus Properties
    [DateTime]$WsusLastUpdated
    [int16]$WsusInstalledCount = -1
    [int16]$WsusInstalledPendingRebootCount = -1
    [int16]$WsusDownloadedCount = -1
    [int16]$WsusFailedCount = -1
    [int16]$WsusNotInstalledCount = -1
    [int16]$WsusTotalNeededCount = -1

    # Methods

    [void]GetADComputer([string]$ADServer)
    {

        $ADComputer = Get-ADComputer -Filter "name -eq '$($this.ComputerName)'" -Server $ADServer -Properties lastLogonTimestamp |
                        Select-Object DNSHostName,Enabled,@{n='LastLogonTime';e={[DateTime]::FromFileTime($_.LastLogonTimeStamp)}}

        $ADComputerCount = ($ADComputer | Measure-Object).Count

        If ( $ADComputerCount -ne 1 )
        {

            throw [System.ArgumentException] "$ADComputer computers returned for name '$($this.ComputerName)'. Expecting one."

        }

        $this.DNSName = $ADComputer.DNSHostName
        $this.ADEnabled = $ADComputer.Enabled
        $this.ADLastLogonTime = $ADComputer.LastLogonTime

    }



    [void]GetIPAndOsDataFromAD($ADServer)
    {
    
        $ADComputer = Get-ADComputer -Filter "name -eq '$($this.ComputerName)'" -Server $ADServer -Properties IPv4Address,OperatingSystem,OperatingSystemVersion,OperatingSystemServicePack

        $ADComputerCount = ($ADComputer | Measure-Object).Count

        If ( $ADComputerCount -ne 1 )
        {

            throw [System.ArgumentException] "$ADComputer computers returned for name '$($this.ComputerName)'. Expecting one."

        }

        $this.IPAddress = $ADComputer.IPv4Address
        $this.OperatingSystem = $ADComputer.OperatingSystem
        $this.OperatingSystemVersion = $ADComputer.OperatingSystemVersion
        $this.OperatingSystemServicePack = $ADComputer.OperatingSystemServicePack

    }

    [void]GetIPAndOsDataFromWsus ([string]$UpdateServer,[int16]$PortNumber)
    {

        $Wsus = Get-WsusServer -Name $UpdateServer -PortNumber $PortNumber -ErrorAction Stop 
        $Target = $Wsus.GetComputerTargetByName($this.DNSName)
        $TargetCount = ($Target | Measure-Object).Count

        If ( $TargetCount -ne 1 )
        {

            throw [System.ArgumentException] "$TargetCount target computers returned for name '$($this.ComputerName)'. Expecting one."

        }


        # Set IP and OS Data
        $this.IPAddress = $Target.IPAddress
        $this.OperatingSystem = $Target.OSDescription
        $this.OperatingSystemVersion = $Target.ClientVersion

    }

    [void]GetWsusComputerSummary ([string]$UpdateServer,[int16]$PortNumber,[string[]]$UpdateClassification)
    {
        
        # Connect to WSUS Server
        $Wsus = Get-WsusServer -Name $UpdateServer -PortNumber $PortNumber -ErrorAction Stop   
       
        # Classifications
        $WsusClassification = $Wsus.GetUpdateClassifications()
        $Classification = $null
    
        ForEach ( $UserClassification in $UpdateClassification )
        {

            If ( ($WsusClassification.Title) -notcontains $UserClassification )
            {
    
                throw [System.ArgumentException] "Update Classification '$UserClassification' cannot be found on Update Server '$($Wsus.Name)'."
    
            }

        }

        If ( $UpdateClassification -ne $null )
        {

            $Classification = $WsusClassification | Where-Object { $UpdateClassification -contains $_.Title }

        }
        else
        {

            $Classification = $WsusClassification

        }

            
        # Set Update Scope
        $UpdateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
        $updatescope.Classifications.AddRange($Classification)
        $UpdateScope.ApprovedStates = New-Object Microsoft.UpdateServices.Administration.ApprovedStates -Property @{value__ = 3}
        #$UpdateScope.ApprovedStates = [Microsoft.UpdateServices.Administration.ApprovedStates]::LatestRevisionApproved
        #$UpdateScope.ApprovedStates += [Microsoft.UpdateServices.Administration.ApprovedStates]::HasStaleUpdateApprovals

        # Get Computer
        $Target = $Wsus.GetComputerTargetByName($this.DNSName)
        $TargetCount = ($Target | Measure-Object).Count

        If ( $TargetCount -ne 1 )
        {

            throw [System.ArgumentException] "$TargetCount target computers returned for name '$($this.ComputerName)'. Expecting one."

        }

        # Set Computer Scope
        $ComputerScope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope
        $ComputerScope.NameIncludes = $this.ComputerName

        # Get Summary Info
        $AllTargetSummaries = $WSUS.GetSummariesPerComputerTarget($updatescope, $computerscope)
        $TargetSummary = $AllTargetSummaries | Where-Object { $_.ComputerTargetId -eq $Target.Id }

        # Set Wsus Data
        $this.WsusLastUpdated = $TargetSummary.LastUpdated
        $this.WsusInstalledCount = $TargetSummary.InstalledCount
        $this.WsusInstalledPendingRebootCount = $TargetSummary.InstalledPendingRebootCount
        $this.WsusDownloadedCount = $TargetSummary.DownloadedCount
        $this.WsusFailedCount = $TargetSummary.FailedCount
        $this.WsusNotInstalledCount = $TargetSummary.NotInstalledCount
        $this.WsusTotalNeededCount = $TargetSummary.InstalledPendingRebootCount + $TargetSummary.DownloadedCount + $TargetSummary.FailedCount + $TargetSummary.NotInstalledCount

    }

    [void]GetWsusComputerMakeAndModel ([string]$UpdateServer,[int16]$PortNumber)
    {

        # Connect to WSUS Server
        $Wsus = Get-WsusServer -Name $UpdateServer -PortNumber $PortNumber -ErrorAction Stop   
    
        $WsusComputer = Get-WsusComputer -UpdateServer $Wsus -NameIncludes $this.DNSName -ErrorAction Stop | Select-Object Make,Model
    
        $this.Manufacturer = $WsusComputer.Make
        $this.Model = $WsusComputer.Model
    
    }

    [void]GetAuditPropertiesFromComputer ()
    {
        
        #########################################
        # Step 1 - Check connection to computer #
        #########################################
        
        if ( Test-Connection $this.ComputerName -Quiet -Count 1 )
        {

            $this.ConnectionSuccessful = 'TRUE'

        }
        else
        {

            $this.ConnectionSuccessful = 'FALSE'
            return

        }

        #####################################
        # Step 2 - Get Information from WMI #
        #####################################

        # Get WMI Properties
        # At this stage always use Select-Object so that as little is held in memory as possible
            
        Try
        {

            ## Win32_ComputerSystem
            $WmiComputerSystem = Get-WmiObject Win32_ComputerSystem -ComputerName $this.ComputerName -ErrorAction Stop | Select-Object Manufacturer,Model

            $this.Manufacturer = $WmiComputerSystem.Manufacturer
            $this.Model = $WmiComputerSystem.Model


            ## Win32_Bios
            $WmiBios = Get-WmiObject Win32_Bios -ComputerName $this.ComputerName -ErrorAction Stop | Select-Object SerialNumber

            $this.SerialNumber = $WmiBios.SerialNumber


            ## Win32_OperatingSystem
            $WmiOsDetails = Get-WmiObject Win32_OperatingSystem -ComputerName $this.ComputerName -ErrorAction Stop |
                Select-Object Caption,Version,ServicePackMajorVersion,OSArchitecture,@{label='LastBootUpTime';expression={$_.ConverttoDateTime($_.lastbootuptime)}}

            $this.OperatingSystem = $WmiOSDetails.Caption
            $this.OperatingSystemVersion = $WmiOSDetails.Version
            $this.OperatingSystemServicePack = $WmiOSDetails.ServicePackMajorVersion
            $this.OSArchitecture = $WmiOSDetails.OSArchitecture
            $this.LastBootUpTime =  $WmiOSDetails.LastBootUpTime.ToString()


            ## SoftwareLicensingProduct
            ## Warning, hardcoded ApplicationId, in tests this always returned OS licence but test was limited to specific environment but run against server OS and client OS
            $WmiFilter = "ApplicationId='55c92734-d682-4d71-983e-d6ec3f16059f' AND PartialProductKey IS NOT NULL"
            $WMISoftwareLicensing = Get-WmiObject SoftwareLicensingProduct -ComputerName $this.ComputerName -Filter $WmiFilter -ErrorAction Stop |
                Select-Object LicenseStatus

            $this.OSLicenseStatus = $WMISoftwareLicensing.LicenseStatus


            # Win32_NetworkAdapterConfiguration (MAC Address)
            # Excludes Servicename Netft, which is 'Failover Cluster Virtual Adapter' 
            $WmiMacAddress = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $this.ComputerName -ErrorAction Stop |
                Where-Object { ($_.IPEnabled -eq $true) -and ($_.ServiceName -ne 'Netft') } | Select-Object -ExpandProperty MacAddress
        
            $this.MacAddress = $WmiMacAddress -join ','

        }
        Catch
        {

            Write-Error $_.Exception.Message -ErrorId $_.FullyQualifiedErrorId

        }


        ##############################################
        # Step 3 - Get information from the Registry #
        ##############################################

        $RRService = $null
        $RemoteRegError = $null
        $Reg = $null
            
        # Check RemoteRegistry Service
        Try
        {
            
            $RRService = Get-Service RemoteRegistry -ComputerName $this.ComputerName -ErrorAction Stop
            
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

                $Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $this.ComputerName)

            }
            catch
            {

                Write-Error $_.Exception.Message -ErrorId $_.FullyQualifiedErrorId

            }
            

            # Get details from Registry
            try
            {

                # Get Powershell Version
                $RegKeyPSVersion = $Reg.OpenSubKey("SOFTWARE\Microsoft\PowerShell\3\PowerShellEngine")
            
                If ( $RegKeyPSVersion -eq $null )
                {
            
                    $RegKeyPSVersion = $Reg.OpenSubKey("SOFTWARE\Microsoft\PowerShell\1\PowerShellEngine")    
            
                }
            
                $this.PSVersion = $RegKeyPSVersion.GetValue("PowerShellVersion")
                $RegKeyPSVersion.Close()
                
                # Get SMB 1
                $RegKeySmb1 = $Reg.OpenSubKey("SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters")
                    
                $Smb1Value = $RegKeySmb1.GetValue("SMB1")
                $RegKeySmb1.Close()
            
                If ( $Smb1Value -ne $null )
                {
            
                    $this.SMB1Enabled = $Smb1Value
            
                }
                Else
                {
            
                    $this.SMB1Enabled = 1
            
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
             
            Write-Warning "Registry service on '$($this.ComputerName)' is not running, its status is '$($RRService.Status)'."

        }
        ElseIf ($RemoteRegError.Status -ne $null)
        {
             
            Write-Error $RemoteRegError.Exception.Message -ErrorId $RemoteRegError.FullyQualifiedErrorId

        }
     
    }
    
}