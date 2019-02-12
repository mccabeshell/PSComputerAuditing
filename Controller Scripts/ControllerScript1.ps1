$ADServer = 'PRECISEDC01'

$ADComputer = Get-ADComputer -Filter * -Server $ADServer -Properties lastLogonTimestamp |
        Select-Object Name,DNSHostName,Enabled,@{n='LastLogonTime';e={[DateTime]::FromFileTime($_.LastLogonTimeStamp)}}

ForEach ( $Computer in $ADComputer )
{

    $ComputerAudit = New-Object -TypeName  ComputerAudit -ArgumentList $Computer.Name

    $ComputerAudit.DNSName = $ADComputer.DNSHostName
    $ComputerAudit.ADEnabled = $ADComputer.Enabled
    $ComputerAudit.ADLastLogonTime = $ADComputer.LastLogonTime

}

