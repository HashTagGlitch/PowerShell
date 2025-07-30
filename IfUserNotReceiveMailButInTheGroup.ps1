Get-UnifiedGroup "PVChina@precisionglobal.com" | Get-UnifiedGroupLinks -LinkType Members
Get-UnifiedGroupLinks -Identity "PVChina@precisionglobal.com" -LinkType Subscribers | Where-Object {$_.Name -eq "jane.xi"}
 --> Absent
Add-UnifiedGroupLinks -Identity "PVChina@precisionglobal.com" -LinkType Subscribers -Links "jane.xi@precisionglobal.com"
Get-UnifiedGroupLinks -Identity "PVChina@precisionglobal.com" -LinkType Subscribers | Where-Object {$_.Name -eq "jane.xi"}

OK