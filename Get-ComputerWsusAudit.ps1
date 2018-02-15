<#
.Synopsis
   Get Windows Update information from a local WSUS server.
.DESCRIPTION
   Connects to the WSUS server. Will return an object listing all the update states for the computer and a count of how many updates are in each state for the computer.
.PARAMETER ComputerName
   The "FullDomainName" of a computer.
   If joined to a domain, this would be a FQDN (Fully Qualified Domain Name), if in a workgroup then just the computer name should be included without the workgroup name.
.PARAMETER UpdateClassification
   The title of an Update Classification. The values input into this parameter are checked against the list of classifications on the update server specified.
   See notes for an example list and further details.
.EXAMPLE
    This will return information about all Critical and Security updates for 'computer1' joined to domain 'mccabeshell.com'.
    Get-ComputerWsusAudit 'computer1.mccabeshell.com' -UpdateClassification 'Critical Updates','Security Updates' -UpdateServer 'WSUSSERVER' -PortNumber '8530'
.EXAMPLE
    This will return information about all Critical and Security updates for computer computer1 if it is in a Workgroup. If it was joined to a domain then an Object not found error would be thrown.
    Get-ComputerWsusAudit 'computer1' -UpdateClassification 'Critical Updates','Security Updates' -UpdateServer 'WSUSSERVER' -PortNumber '8530'
.NOTES
   List of Update Classifications. It is possible that these classifications may vary with different versions of WSUS and over time.

   Applications
   Critical Updates
   Definition Updates
   Drivers
   Feature Packs
   Security Updates
   Service Packs
   Tools
   Update Rollups
   Updates
   Upgrades

   Due to the fact you need to know the exact title, a new cmdlet may be added to the PSComputerAudit repository to get UpdateClassifications.
   This function already does this to get the list to compare to. It uses the method GetUpdateClassifications, returning objects of "Microsoft.UpdateServices.Internal.BaseApi.UpdateClassification".
   See link for more details about this method.
.LINK
    https://msdn.microsoft.com/en-us/library/windows/desktop/ms747071(v=vs.85).aspx
#>

Function Get-ComputerWsusAudit
{

    [CmdletBinding()]
    Param
    (

        [Parameter( Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true,
            Position=0)]
        [ValidateNotNullOrEmpty()]
        [string[]]$ComputerName,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$UpdateClassification,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$UpdateServer,
        
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$PortNumber
        
    )

    BEGIN
    {

        Class ComputerWsusAudit
        {
            [string]$ComputerName
            [DateTime]$LastUpdated
            [int16]$InstalledCount
            [int16]$InstalledPendingRebootCount
            [int16]$DownloadedCount
            [int16]$FailedCount
            [int16]$NotInstalledCount
        }

        ###################
        # Get WSUS Server #
        ###################
        try
        {
        
            $Wsus = Get-WsusServer -Name $UpdateServer -PortNumber $PortNumber -ErrorAction Stop   

        }
        catch
        {
            
            throw "Unable to connect to WSUS server '$UpdateServer' on port '$PortNumber'."
                    
        }
        
        
        #################################
        # Get and Check Classifications #
        #################################
        try
        {
        
            $WsusClassification = $Wsus.GetUpdateClassifications()
        
            If ( $UpdateClassification -ne $null )
            {
    
                ForEach ( $UserClassification in $UpdateClassification )
                {
                    If ( ($WsusClassification.Title) -notcontains $UserClassification )
                    {
    
                        throw "Update Classification '$UserClassification' cannot be found on Update Server '$($Wsus.Name)'."
    
                    }
                }
    
                $Classification = $WsusClassification | Where-Object { $UpdateClassification -contains $_.Title }
    
            }

        }
        catch
        {
        
            throw $_

        }
        
        ##################################
        # Set Update and Computer Scopes #
        ##################################
        try
        {
            
            # Set Update Scope
            $UpdateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
            $UpdateScope.ApprovedStates = [Microsoft.UpdateServices.Administration.ApprovedStates]::LatestRevisionApproved
            $updatescope.IncludedInstallationStates = [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::All
            $updatescope.Classifications.AddRange($Classification)

            # Set Computer Scope (get all computers)
            $ComputerScope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope
        
        }
        catch
        {
            
            throw $_
        
        }


        #######################################
        # Fetch all target computer summaries #
        #######################################
        try
        {
            
            $AllComputerTargetSummaries = $WSUS.GetSummariesPerComputerTarget($updatescope, $computerscope)
        
        }
        catch
        {
            
            throw $_
        
        }

    }

    PROCESS
    {

        foreach ( $Computer in $ComputerName )
        {
            
            try
            {

                $ComputerTarget = $Wsus.GetComputerTargetByName($Computer)

            }
            catch [System.Management.Automation.MethodInvocationException]
            {

                Write-Error $_.Exception.Message
                continue

            }
            
            
            $ComputerTargetSummary = $AllComputerTargetSummaries | Where-Object { $_.ComputerTargetId -eq $ComputerTarget.Id }
            

            $ComputerWsusAudit = New-Object ComputerWsusAudit
            
                
            $ComputerWsusAudit.ComputerName = $ComputerTarget.FullDomainName
            $ComputerWsusAudit.DownloadedCount = $ComputerTargetSummary.DownloadedCount
            $ComputerWsusAudit.FailedCount = $ComputerTargetSummary.FailedCount
            $ComputerWsusAudit.InstalledCount = $ComputerTargetSummary.InstalledCount
            $ComputerWsusAudit.InstalledPendingRebootCount = $ComputerTargetSummary.InstalledPendingRebootCount
            $ComputerWsusAudit.LastUpdated = $ComputerTargetSummary.LastUpdated
            $ComputerWsusAudit.NotInstalledCount = $ComputerTargetSummary.NotInstalledCount
        
            Write-Output $ComputerWsusAudit
            Clear-Variable ComputerWsusAudit

        }

    }

    END
    {

    }

}