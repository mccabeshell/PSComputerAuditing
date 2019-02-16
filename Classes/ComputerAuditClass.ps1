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
    
    ## Properties Generally from Remote/local Connection
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

    # Constructor
    ComputerAudit ([string]$ComputerName)
    {
        $this.ComputerName = $ComputerName
    }

}