function Get-CitrixRSOP
{
<#
.Synopsis
Retrieves the Citrix group policy objects applied to a computer
.DESCRIPTION
This function retrieves the Citrix group policy objects that have been applied to the target
computer. User policies will be retrieved if a user is logged on. For this function to work
the Citrix Group Policy Management tools must be installed
.EXAMPLE
Get-CitrixRSOP -ComputerName "TargetPC01"
.EXAMPLE
Get-CitrixRSOP -ComputerName "TargetPC01" -Protocol DCOM
.PARAMETER ComputerName
The computer to retrieve the policies from
.PARAMETER Protocol
The protocol used to retrieve the WMI objects from the remote computer. Defaults to CIM, but
DCOM can be specified as well.
.PARAMETER SessionId
The user session ID to get user policies for. For XenDesktop this will always be 1 (the default).
#>
  [CmdletBinding()]

  param(
    [Parameter(Mandatory = $true,ValueFromPipeline = $True,ValueFromPipelineByPropertyName = $True)]
    [ValidateNotNullOrEmpty()]
    [string[]]$ComputerName,
    [Parameter(Mandatory = $false,ValueFromPipeline = $True,ValueFromPipelineByPropertyName = $True)]
    [ValidateSet("CIM","DCOM")]
    [string]$Protocol = "CIM",
    [Parameter(Mandatory = $false,ValueFromPipeline = $True,ValueFromPipelineByPropertyName = $True)]
    [int]$SessionId = 1
  )

  begin
  {
    Set-StrictMode -Version Latest
    $ErrorActionPreference = "Stop"
    try
    {
      $null = [System.Reflection.Assembly]::LoadwithPartialName("Citrix.GroupPolicy.DataModel")
      $null = [System.Reflection.Assembly]::LoadwithPartialName("Citrix.GroupPolicy.Utilities")
      if (-not ([System.Management.Automation.PSTypeName]'Citrix.GroupPolicy.DataModel.CGPReportTranslator').Type)
      {
        throw "Citrix Group Policy tools not installed"
      }
      
      if (Test-Path "REGISTRY::HKEY_CLASSES_ROOT\CLSID\{54C5637D-BAC7-4C38-A2FA-E314971F6090}")
      {
        $regprops = Get-ItemProperty "REGISTRY::HKEY_CLASSES_ROOT\CLSID\{54C5637D-BAC7-4C38-A2FA-E314971F6090}\InprocServer32"
        $dll = Get-Item $regprops.'(default)'
        Write-Verbose "Group policy tools version: $($dll.VersionInfo.FileVersion)"
      }
    }
    catch
    {
      Write-Output $_.Exception.Message
      break
    }
  }
  process
  {
    foreach ($Computer in $ComputerName)
    {
      try
      {
        Write-Verbose "Retrieving policy for $Computer"
        $test = Test-WSMan -ComputerName $Computer
        
        switch ($Protocol)
        {
          "CIM"
          {
            Write-Verbose "Connecting via CIM"
            $CitrixRsopProviderClass = Get-CimClass -Namespace root\rsop -ClassName CitrixRsopProviderClass
            $options = New-CimSessionOption -MaxEnvelopeSizeKB 2048
            $session = New-CimSession -ComputerName $Computer -SessionOption $options
            $wmiC = Invoke-CimMethod -CimClass $CitrixRsopProviderClass -MethodName GetRsopRawData -Arguments @{ 'IsComputer' = $true; 'username' = "" } -CimSession $session
            $wmiU = Invoke-CimMethod -CimClass $CitrixRsopProviderClass -MethodName GetRsopRawDataForSession -Arguments @{ 'sessionId' = $SessionId } -CimSession $session
            $user = (Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $Computer).UserName
          }
          "DCOM"
          {
            Write-Verbose "Connecting via DCOM"
            $wmiC = Invoke-WmiMethod -Class CitrixRsopProviderClass -Name GetRsopRawData -Namespace root\rsop -ComputerName $Computer -ArgumentList $true,""
            $wmiU = Invoke-WmiMethod -Class CitrixRsopProviderClass -Name GetRsopRawDataForSession -Namespace root\rsop -ComputerName $Computer -ArgumentList $SessionId
            $user = (Get-WmiObject -ClassName Win32_ComputerSystem -ComputerName $Computer).UserName
          }

        }
        $streamC = New-Object System.IO.MemoryStream (,$wmiC.data)
        $streamU = New-Object System.IO.MemoryStream (,$wmiU.data)
        $resultC = [Citrix.GroupPolicy.DataModel.CGPReportTranslator]::CreateRsopPoliciesOnlyReport($streamC,[Citrix.GroupPolicy.Utilities.Phase]::Computer)
        $resultU = [Citrix.GroupPolicy.DataModel.CGPReportTranslator]::CreateRsopPoliciesOnlyReport($streamU,[Citrix.GroupPolicy.Utilities.Phase]::Computer)
        $streamC.Dispose()
        $streamU.Dispose()
        $props = [ordered]@{ 'ComputerName' = $Computer;
                             'UserName' = $user;
                             'ComputerPolicies' = $resultC;
                             'UserPolicies' = $resultU;
                           }
        Write-Output (New-Object -TypeName PSObject -Property $props)
        
      }
      catch
      {
        Write-Output $_.Exception.Message
      }
    }
  }
  end
  { Write-Verbose "Finished Processing" }
}
