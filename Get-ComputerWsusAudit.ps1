<#
.Synopsis
   Get Windows Update information from a local WSUS server.
.DESCRIPTION
   Connects to the WSUS server. Will return an object listing all the update states for the computer and a count of how many updates are in each state for the computer.
.PARAMETER ComputerName
   The name of a computer. By default this will be the FullDomainName or FQDN unless the SwitchOffExactNameMatch switch parameter is used.
.PARAMETER SwitchOffExactNameMatch
   If used, this will return any Wsus computer objects whose name conatins the string in the ComputerName parameter. This can return multiple objects per ComputerName.
   If not used, the cmdlet will only return an exact match and therefore only one object can be output per ComputerName.
.PARAMETER UpdateClassification
   The title of an Update Classification. Only updates in these classifications are returned in the results.
   If left as default $null, updates for all classifications are included.
   The values input into this parameter are checked against the list of classifications on the update server specified.
   To get a list of Update Classifications, use the Powershell cmdlet Get-WsusClassification.
.PARAMETER IncludeUnapprovedUpdates
   Updates with a status of either LatestRevisionApproved or HasStaleUpdateApprovals are always included.
   Using this switch will also include updates that have not been approved or declined.
   Declined updates are never included in the results of this function.
.EXAMPLE
    This will return information about all Critical and Security updates for 'computer1' joined to domain 'mccabeshell.com'.
    Get-ComputerWsusAudit 'computer1.mccabeshell.com' -UpdateClassification 'Critical Updates','Security Updates' -UpdateServer 'WSUSSERVER' -PortNumber '8530'
.EXAMPLE
    This will return information about all Critical and Security updates for computer computer1 if it is in a Workgroup. If it was joined to a domain then an Object not found error would be thrown.
    Get-ComputerWsusAudit 'computer1' -UpdateClassification 'Critical Updates','Security Updates' -UpdateServer 'WSUSSERVER' -PortNumber '8530'
.LINK
    https://msdn.microsoft.com/en-us/library/windows/desktop/ms747071(v=vs.85).aspx
#>

Function Get-ComputerWsusAudit
{

    [CmdletBinding()]
    Param
    (

        [Parameter(Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true,
            Position=0)]
        [ValidateNotNullOrEmpty()]
        [string[]]$ComputerName,

        [switch]$SwitchOffExactNameMatch,

        [string[]]$UpdateClassification = $null,

        [switch]$IncludeUnapprovedUpdates,

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
            [int16]$TotalNeededCount
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
            
            throw $_
                    
        }
        
        
        ##################################
        # Get then Check Classifications #
        ##################################
        try
        {
        
            $WsusClassification = $Wsus.GetUpdateClassifications()

        }
        catch
        {
        
            throw $_

        }

    
        If ( $UpdateClassification -ne $null )
        {

            ForEach ( $UserClassification in $UpdateClassification )
            {
                If ( ($WsusClassification.Title) -notcontains $UserClassification )
                {
    
                    throw [System.ArgumentException] "Update Classification '$UserClassification' cannot be found on Update Server '$($Wsus.Name)'."
    
                }

                $Classification = $WsusClassification | Where-Object { $UpdateClassification -contains $_.Title }

            }

        }
        else
        {

            $Classification = $WsusClassification

        }
        
        ##################################
        # Set Update and Computer Scopes #
        ##################################
        try
        {
            
            # Set Update Scope
            $UpdateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
            $updatescope.Classifications.AddRange($Classification)
            $updatescope.IncludedInstallationStates = [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::All
            $UpdateScope.ApprovedStates = [Microsoft.UpdateServices.Administration.ApprovedStates]::LatestRevisionApproved
            $UpdateScope.ApprovedStates += [Microsoft.UpdateServices.Administration.ApprovedStates]::HasStaleUpdateApprovals

            If ( $IncludeUnapprovedUpdates )
            {

                $UpdateScope.ApprovedStates += [Microsoft.UpdateServices.Administration.ApprovedStates]::NotApproved

            }
        
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

            Write-Verbose "Fetching update summary for target '$Computer' on Update Server '$UpdateServer'"

            ####################################
            # Fetch  target computer summaries #
            ####################################
            try
            {

                $ComputerScope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope
                $ComputerScope.NameIncludes = $Computer              
                $ComputerTargetSummaries = $WSUS.GetSummariesPerComputerTarget($updatescope, $computerscope)
            
            }
            catch
            {
                
                Write-Error $_
                continue
            
            }
            
            # Filter exact match on name
            if ( -not $SwitchOffExactNameMatch )
            {

                try
                {

                    $TargetName = $Wsus.GetComputerTargetByName($Computer)
                    $ComputerTargetSummaries = $ComputerTargetSummaries | Where-Object { $_.ComputerTargetId -eq $TargetName.Id }
                    
                }
                catch
                {
                    
                    Write-Error $_
                    continue
                
                }

            }

            #################################################
            # Create and output ComputerWsusAudit object(s) #
            #################################################
            ForEach ( $Target in $ComputerTargetSummaries )
            {

                if ( $SwitchOffExactNameMatch )
                {
    
                    try
                    {
                     
                        $TargetName = $Wsus.GetComputerTarget($Target.ComputerTargetId)
                    
                    }
                    catch
                    {

                        Write-Error $_
                        continue

                    }
                }

                $ComputerWsusAudit = New-Object ComputerWsusAudit
            
                $ComputerWsusAudit.ComputerName = $TargetName.FullDomainName
                $ComputerWsusAudit.LastUpdated = $Target.LastUpdated
                $ComputerWsusAudit.InstalledCount = $Target.InstalledCount
                $ComputerWsusAudit.InstalledPendingRebootCount = $Target.InstalledPendingRebootCount
                $ComputerWsusAudit.DownloadedCount = $Target.DownloadedCount
                $ComputerWsusAudit.FailedCount = $Target.FailedCount
                $ComputerWsusAudit.NotInstalledCount = $Target.NotInstalledCount
                $ComputerWsusAudit.TotalNeededCount = $Target.InstalledPendingRebootCount + $Target.DownloadedCount + $Target.FailedCount + $Target.NotInstalledCount
    
                Write-Output $ComputerWsusAudit
                Clear-Variable ComputerWsusAudit

            }


        }

    }

    END
    {

    }

}