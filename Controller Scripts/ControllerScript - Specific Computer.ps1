$ADServer = 'PRECISEDC01'

$ComputerAudit = New-Object ComputerAudit -ArgumentList $env:COMPUTERNAME

$ADResults = Get-CAActiveDirectoryProperties -ComputerName $ComputerAudit.ComputerName -ADServer $ADServer -Verbose

$ComputerAudit.DNSName = $ADResults.DNSHostName
$ComputerAudit.ADEnabled = $ADResults.Enabled
$ComputerAudit.ADLastLogonTime = $ADResults.LastLogonTime
$ComputerAudit.IPAddress = $ADResults.IPv4Address
$ComputerAudit.OperatingSystem = $ADResults.OperatingSystem
$ComputerAudit.OperatingSystemVersion = $ADResults.OperatingSystemVersion
$ComputerAudit.OperatingSystemServicePack = $ADResults.OperatingSystemServicePack


$ComputerAudit