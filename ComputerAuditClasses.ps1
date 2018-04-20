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
    
    
    ## Properties Generally from Remote Connection
    [string]$MacAddress
    [string]$LastBootUpTime
    [string]$Manufacturer
    [string]$Model
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
        $updatescope.IncludedInstallationStates = [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::All
        $UpdateScope.ApprovedStates = [Microsoft.UpdateServices.Administration.ApprovedStates]::LatestRevisionApproved
        $UpdateScope.ApprovedStates += [Microsoft.UpdateServices.Administration.ApprovedStates]::HasStaleUpdateApprovals

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
    
}