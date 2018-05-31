<#
.Synopsis
   Get Windows Update counts from a local WSUS server.
.DESCRIPTION
   Connects to a WSUS server returning all Wsus computers and counts of updates with of each status, such as Installed, Downloaded, Failed, etc. as well as a calculated count of 
   "TotalNeededCount", a sum of all the statuses except "Installed".
.PARAMETER UpdateServer
   The name of a computer running the Update Services role.
.PARAMETER PortNumber
   The Port used by the Wsus server. On a default installation this would be 8530, which is also the default for this parameter.
.PARAMETER UpdateClassification
   The title of an Update Classification. Only updates in these classifications are returned in the results.
   If left as default $null, updates for all classifications are included.
   The values input into this parameter are checked against the list of classifications on the update server specified.
   To get a list of Update Classifications, use the Powershell cmdlet Get-WsusClassification.
.EXAMPLE
    This will return information about all Critical and Security updates for 'computer1' joined to domain 'mccabeshell.com'.
    Get-WsusComputerSummary  -UpdateServer 'WSUSSERVER' -PortNumber '8530' -UpdateClassification 'Critical Updates','Security Updates'
.NOTES
    This cmdlet is designed as a simple and quick method to list all computers in Wsus and which need updates. Future development may add parameters to allow filtering, such as
    returning only those needing updates.
.LINK
    https://msdn.microsoft.com/en-us/library/windows/desktop/ms747071(v=vs.85).aspx
#>

function Get-WsusComputerSummary
{

    [CmdletBinding(SupportsShouldProcess=$true, 
                   PositionalBinding=$false)]

    Param
    (
    
        # Params
        [Parameter]
        [string]$UpdateServer,
        [string]$PortNumber = 8530,
        [string[]]$UpdateClassification
    
    )

    Class ComputerWsusAudit
    {
        [string]$ComputerName
        [DateTime]$LastUpdated
        [int16]$InstalledCount
        [int16]$InstalledPendingRebootCount
        [int16]$DownloadedCount
        [int16]$FailedCount
        [int16]$NotInstalledCount
        [int16]$TotalNeededCount
    }

    #BEGIN BLOCK

    # Get WSUS Server
    $Wsus = Get-WsusServer -Name $UpdateServer -PortNumber $PortNumber

    # Get Classifications - Do a check against geuine names (use a loop)...
    $WsusClassification = $Wsus.GetUpdateClassifications()

    ForEach ( $UserClassification in $UpdateClassification )
    {
        If ( ($WsusClassification.Title) -notcontains $UserClassification )
        {

            Write-Error "Update Classification '$UserClassification' cannot be found on Update Server '$($Wsus.Name)'."
            # Terminate here

        }
    }

    $Classification = $WsusClassification | Where-Object { $UpdateClassification -contains $_.Title }

    # Set Update Scope
    $UpdateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
    $UpdateScope.ApprovedStates = [Microsoft.UpdateServices.Administration.ApprovedStates]::LatestRevisionApproved
    $updatescope.IncludedInstallationStates = [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::All
    $updatescope.Classifications.AddRange($Classification)

    $ComputerScope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope

    $AllComputerTargetSummaries = $WSUS.GetSummariesPerComputerTarget($updatescope, $computerscope)


    foreach ( $TargetSummary in $AllComputerTargetSummaries )
    {

        $ComputerTarget = $Wsus.GetComputerTarget($TargetSummary.ComputerTargetId)

        $ComputerWsusAudit = New-Object ComputerWsusAudit
        
        $ComputerWsusAudit.ComputerName = $ComputerTarget.FullDomainName
        $ComputerWsusAudit.LastUpdated = $TargetSummary.LastUpdated
        $ComputerWsusAudit.InstalledCount = $TargetSummary.InstalledCount
        $ComputerWsusAudit.InstalledPendingRebootCount = $TargetSummary.InstalledPendingRebootCount
        $ComputerWsusAudit.DownloadedCount = $TargetSummary.DownloadedCount
        $ComputerWsusAudit.FailedCount = $TargetSummary.FailedCount
        $ComputerWsusAudit.NotInstalledCount = $TargetSummary.NotInstalledCount
        $ComputerWsusAudit.TotalNeededCount = $TargetSummary.InstalledPendingRebootCount + $TargetSummary.DownloadedCount + $TargetSummary.FailedCount + $TargetSummary.NotInstalledCount

        Write-Output $ComputerWsusAudit
        Clear-Variable ComputerWsusAudit

    }

}