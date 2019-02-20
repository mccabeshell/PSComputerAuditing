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
.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
   The functionality that best describes this cmdlet
#>
function Get-CARegistryData
{
    [CmdletBinding(SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
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

         $Reg = $null

         try
         {
            
            # Open the registry
            $Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Computer)


            # Get Powershell Version
            $RegKeyPSVersion = $Reg.OpenSubKey("SOFTWARE\Microsoft\PowerShell\3\PowerShellEngine")
            
            If ( $null -eq $RegKeyPSVersion )
            {
            
               $RegKeyPSVersion = $Reg.OpenSubKey("SOFTWARE\Microsoft\PowerShell\1\PowerShellEngine")    
            
            }
            
            $PSVersion = $RegKeyPSVersion.GetValue("PowerShellVersion")
            $RegKeyPSVersion.Close()
               
            # Get SMB 1
            $RegKeySmb1 = $Reg.OpenSubKey("SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters")
                     
            $Smb1Value = $RegKeySmb1.GetValue("SMB1")
            $RegKeySmb1.Close()
            
            If ( $null -ne $Smb1Value )
            {
            
               $SMB1Enabled = $Smb1Value
            
            }
            Else
            {
            
               $SMB1Enabled = 1
            
            }

            $OutputParams = @{

               ComputerName = $Computer;
               PSVersion = $PSVersion;
               SMB1Enabled = $SMB1Enabled;


            }

            $OutputObject = New-Object -TypeName psobject -ArgumentList $OutputParams

            Write-Output $OutputObject

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

      }

    }
    End
    {
    }
}