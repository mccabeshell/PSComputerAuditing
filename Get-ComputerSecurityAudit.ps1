Function Get-ComputerSecurityAudit
{

    [CmdletBinding()]
    Param
    (
       [Parameter(Position=0)]
       [string]$ComputerName = 'localhost'

    )

    Begin
    {

        Class ComputerAudit
        {
            [string]$ComputerName
            [string]$LastBootUpTime = ''
            [string]$OperatingSystem = ''
            [string]$PSVersion = ''
            #-1 = Information Unavailable
            [sbyte]$SMB1Enabled = -1
            [sbyte]$LicenseStatus = -1


        }

    }

    Process
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

                Write-Error "Cannot connect to computer '$Computer'."
                Continue

            }
            
            #####################################
            # Step 2 - Get Information from WMI #
            #####################################

            # Get Licence
            Try
            {

                $ComputerAudit.LicenseStatus = Get-WmiObject SoftwareLicensingProduct -ComputerName $Computer -ErrorAction Stop |
                    Where-Object {$_.PartialProductKey} |
                        Select-Object -ExpandProperty LicenseStatus

            }
            Catch
            {

                $ComputerAudit.LicenseStatus = -1

            }

            #Get OS Details
            try
            {
            
                $WmiOsDetails = Get-WmiObject Win32_OperatingSystem -ComputerName $Computer -ErrorAction Stop |
                    Select-Object Version, @{label='LastBootUpTime';expression={$_.ConverttoDateTime($_.lastbootuptime)}}


                $ComputerAudit.OperatingSystem =  $WmiOSDetails.Version
                $ComputerAudit.LastBootUpTime =  $WmiOSDetails.LastBootUpTime

            }
            catch
            {
             
                $ComputerAudit.OperatingSystem = ''
                $ComputerAudit.LastBootUpTime = ''

            }
                
            ##############################################
            # Step 3 - Get information from the Registry #
            ##############################################
            
            # Check RemoteRegistry Service
            Try
            {
            
                $RRService = Get-Service RemoteRegistry -ComputerName $Computer -ErrorAction Stop
            
                If ( $RRService.Status -ne 'Running' )
                {
                
                    Write-Warning "Remote registry service on computer '$Computer' is not running."
                
                }

            }
            Catch
            {

                Write-Error "Unable to get RemoteRegistry service for computer '$Computer'."

            }

            
            # Open the registry
            Try
            {

                $Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Computer)

            }
            Catch
            {

                Throw $_

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
            
                $ComputerAudit.PSVersion = $RegKeyPSVersion.GetValue("PowerShellVersion")
            
                
                # Get SMB 1
                $RegKeySmb1 = $Reg.OpenSubKey("SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters")
            
                If ( $RegKeySmb1 -ne $null )
                {
            
                    $ComputerAudit.SMB1Enabled = $RegKeySmb1.GetValue("SMB1")
            
                }
                Else
                {
            
                    $ComputerAudit.SMB1Enabled = -1
            
                }

            }#EndOfTry

            Catch
            {
            
                Throw $_
            
            }


            #################################
            # Output a ComputerAudit object #
            #################################
        
            Write-Output $ComputerAudit
    
            Clear-Variable ComputerAudit
         
        }#EndOfProcessForEach

    }

    end
    {

        #No action in end block

    }

}#EndOfFunction
