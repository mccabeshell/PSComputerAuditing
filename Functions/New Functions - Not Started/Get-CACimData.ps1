<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   General notes
#>
function Get-CACimData
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  HelpUri = 'http://www.microsoft.com/',
                  ConfirmImpact='Medium')]

   #Requires -Version 3.0

   Param
   (
      # Param1 help description
      [Parameter(Mandatory=$true, 
                  ValueFromPipeline=$true,
                  ValueFromPipelineByPropertyName=$true)]
      [ValidateNotNullOrEmpty()]
      [string[]]$ComputerName

   )

    Begin
    {
    }
    Process
    {

      foreach ( $Computer in $ComputerName )
      {

         if ($pscmdlet.ShouldProcess("$Computer", "Getting CIM data"))
         {

            Try
            {

                  ## Win32_ComputerSystem
                  $CimComputerSystem = Get-CimInstance Win32_ComputerSystem -ComputerName $Computer -ErrorAction Stop

                  ## Win32_Bios
                  $CimBios = Get-CimInstance Win32_Bios -ComputerName $Computer -ErrorAction Stop

                  ## Win32_OperatingSystem
                  $CimOsDetails = Get-CimInstance Win32_OperatingSystem -ComputerName $Computer -ErrorAction Stop

                  ## SoftwareLicensingProduct
                  ## Warning, hardcoded ApplicationId, in tests this always returned OS licence but test was limited to specific environment but run against server OS and client OS
                  $CimFilter = "ApplicationId='55c92734-d682-4d71-983e-d6ec3f16059f' AND PartialProductKey IS NOT NULL"
                  $CimSoftwareLicensing = Get-CimInstance SoftwareLicensingProduct -ComputerName $Computer -Filter $CimFilter -ErrorAction Stop

                  # Win32_NetworkAdapterConfiguration (MAC Address)
                  # Excludes Servicename Netft, which is 'Failover Cluster Virtual Adapter' 
                  $CimMacAddress = Get-CimInstance Win32_NetworkAdapterConfiguration -ComputerName $Computer -ErrorAction Stop |
                     Where-Object { ($_.IPEnabled -eq $true) -and ($_.ServiceName -ne 'Netft') }

            }
            Catch
            {

                  Write-Error $_.Exception.Message -ErrorId $_.FullyQualifiedErrorId
                  continue

            }

            $OutputParams = @{

               ComputerName = $Computer;
               Manufacturer = $CimComputerSystem.Manufacturer
               Model = $CimComputerSystem.Model;
               SerialNumber = $CimBios.SerialNumber;
               OperatingSystem = $CimOsDetails.Caption;
               OperatingSystemVersion = $CimOsDetails.Version;
               OperatingSystemServicePack = $CimOsDetails.ServicePackMajorVersion;
               OSArchitecture = $CimOsDetails.OSArchitecture;
               LastBootUpTime = $CimOsDetails.LastBootUpTime;
               OSLicenseStatus = $CimSoftwareLicensing.LicenseStatus
               MacAddress = $CimMacAddress.MACAddress -join ','

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