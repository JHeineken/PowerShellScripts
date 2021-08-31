#Get list of all non-server AD computers with Operating System, Name, and Last Logon Date; generate excel file
Get-ADComputer -Filter  {OperatingSystem -notLike '*SERVER*' } -Properties lastlogondate,operatingsystem | select name,lastlogondate,operatingsystem |
    export-csv c:\users\jheineken ComputerList.csv -notypeinformation 

#Get list of AD users & their last logon time, generate excel file
Get-ADuser -Filter * Properties LastLogonTimeStamp | select-object Name,@{Name="LastLogonTime"; Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}} |
    export-csv c:\users.csv –notypeinformation

Test-WSMan

winrm get winrm/config/client
winrm get winrm/config/service

winrm enumerate winrm/config/listener

$credential = Get-Credential #load credentials via terminal#

Enter-PSSession -Computername "machine name" -Credential $credential #connection to single machine through the PS terminal#

Test-WSMan FTN-TS-1525 -Authentication Negotiate -Credential $credential #Confirm you can remotely connect to machine with Powershell/WinRM#

Get-NetTCPConnection -Localport 5985 #Is THIS machine listening on the specified port#

Test-NetConnection -Computername 10.100.1.101 -Port 5985 #Is the REMOTE machine listening on the specified port#

Get-CimInstance -Classname Win32_ComputerSystem -Computername "ftn-winrm-test" | Format-Table #returns Name, primary owner, domain, RAM, Model, Manufacturer, PSComputername#