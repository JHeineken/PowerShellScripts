$computerList = Get-ADComputer -Filter * -Properties * 


foreach($comp in $computerList)
{
    Get-CimInstance -Classname Win32_ComputerSystem -Computername $comp.Name | Format-Table

}


#TODO:
#Look into WinRM Firewall permissions
#Error "WinRM client cannot process the request because the server name cannot be resolved"