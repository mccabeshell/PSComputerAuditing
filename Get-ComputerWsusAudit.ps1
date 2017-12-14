Function Get-ComputerWsusAudit
{

    Param
    (

        [Parameter( Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true,
            Position=0)]
        [ValidateNotNullOrEmpty()]
        [string[]]$ComputerName,

        [string[]]$UpdateClassification = $null,

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