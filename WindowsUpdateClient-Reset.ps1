# Reset Windows Update
# Simple script to fix majority of issues with Windows Updates on a client computer

$SDdir = "$env:systemroot\SoftwareDistribution"
$WuServices = 'wuauserv','bits','appidsvc','cryptsvc'

Get-Service $WuServices | Stop-Service -Verbose
If ( Test-Path "$SDdir.old" ) { Remove-Item "$SDdir.old" -Recurse -Force -Confirm:$false -Verbose }
Remove-Item $SDdir -Force -Recurse -Confirm:$false -Verbose


get-item "$env:ALLUSERSPROFILE\Application Data\Microsoft\Network\Downloader\qmgr*.dat" | Remove-Item -Force -Verbose
Set-Location "$env:windir\system32"
regsvr32.exe atl.dll
regsvr32.exe urlmon.dll
regsvr32.exe mshtml.dll
regsvr32.exe shdocvw.dll
regsvr32.exe browseui.dll
regsvr32.exe jscript.dll
regsvr32.exe vbscript.dll
regsvr32.exe scrrun.dll
regsvr32.exe msxml.dll
regsvr32.exe msxml3.dll
regsvr32.exe msxml6.dll
regsvr32.exe actxprxy.dll
regsvr32.exe softpub.dll
regsvr32.exe wintrust.dll
regsvr32.exe dssenh.dll
regsvr32.exe rsaenh.dll
regsvr32.exe gpkcsp.dll
regsvr32.exe sccbase.dll
regsvr32.exe slbcsp.dll
regsvr32.exe cryptdlg.dll
regsvr32.exe oleaut32.dll
regsvr32.exe ole32.dll
regsvr32.exe shell32.dll
regsvr32.exe initpki.dll
regsvr32.exe wuapi.dll
regsvr32.exe wuaueng.dll
regsvr32.exe wuaueng1.dll
regsvr32.exe wucltui.dll
regsvr32.exe wups.dll
regsvr32.exe wups2.dll
regsvr32.exe wuweb.dll
regsvr32.exe qmgr.dll
regsvr32.exe qmgrprxy.dll
regsvr32.exe wucltux.dll
regsvr32.exe muweb.dll
regsvr32.exe wuwebv.dll

netsh winsock reset

Start-Sleep 5

Restart-Computer -Force -Verbose
