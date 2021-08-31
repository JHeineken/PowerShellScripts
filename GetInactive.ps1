$DaysInactive = 90

$time = (Get-Date).Adddays(-($DaysInactive))

Get-ADComputer -Filter {LastLogonTimeStamp -lt $time} -Properties * | Format-Table * -Wrap -AutoSize