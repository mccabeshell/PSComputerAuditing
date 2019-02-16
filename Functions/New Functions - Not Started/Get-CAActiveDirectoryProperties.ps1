<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   General notes
#>
function Get-CAActiveDirectoryProperties
{
    [CmdletBinding(SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Low')]


   #Requires -Modules ActiveDirectory

    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$ComputerName,

        # Param2 help description
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$ADServer

    )

    Begin
    {

      if ( -not ( Test-Connection $ADServer -Count 1 -Quiet ) ) 
      {

         throw [System.Net.NetworkInformation.PingException]::new("Cannot connect to AD server '$ADServer'")
      
      }

    }
    Process
    {

      ForEach ( $Computer in $ComputerName )
      {

         if ($pscmdlet.ShouldProcess("$Computer", "Get Active Directory record"))
         {

            try
            {

                $ADResults = Get-ADComputer -Filter "name -eq '$Computer'" -Server $ADServer -Properties DNSHostName,IPv4Address,OperatingSystem,OperatingSystemVersion,OperatingSystemServicePack,Enabled,lastLogonTimestamp -ErrorAction Stop

                $ADResultCount = ($ADResults | Measure-Object).Count

                if ( $ADResultCount -ne 1 )
                {

                    throw "Found $ADResultCount computers with name '$Computer' in Active Directory, expecting 1."

                }

            }
            catch
            {

                Write-Error $_
                continue

            }

            $OutputParams = @{

               ComputerName = $ADResults.Name;
               Enabled = $ADResults.Enabled;
               LastLogonTime = [DateTime]::FromFileTime($ADResults.LastLogonTimeStamp)
               DNSHostName = $ADResults.DNSHostName;
               IPv4Address = $ADResults.IPv4Address;
               OperatingSystem = $ADResults.OperatingSystem;
               OperatingSystemVersion = $ADResults.OperatingSystemVersion;
               OperatingSystemServicePack = $ADResults.OperatingSystemServicePack;

            }

            $OutputObject = New-Object -TypeName psobject -ArgumentList $OutputParams

            Write-Output $OutputObject

         }

      }

    }
    End
    {
    }
}