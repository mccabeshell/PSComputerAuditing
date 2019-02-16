#This needs work, currently Get-CAActiveDirectoryProperties is designed to return only one computer

$ADServer = ''

Get-CAActiveDirectoryProperties -ComputerName * -ADServer $ADServer -Verbose | ForEach-Object {

    $ComputerAudit = New-Object ComputerAudit -ArgumentList $env:COMPUTERNAME

    $ComputerAudit.DNSName = $ADResults.DNSHostName
    $ComputerAudit.ADEnabled = $ADResults.Enabled
    $ComputerAudit.ADLastLogonTime = $ADResults.LastLogonTime
    $ComputerAudit.IPAddress = $ADResults.IPv4Address
    $ComputerAudit.OperatingSystem = $ADResults.OperatingSystem
    $ComputerAudit.OperatingSystemVersion = $ADResults.OperatingSystemVersion
    $ComputerAudit.OperatingSystemServicePack = $ADResults.OperatingSystemServicePack

    $ComputerAudit

}