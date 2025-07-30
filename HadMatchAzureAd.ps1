*** Powershel 5.1 *** On my Computer
Install-Module AzureAD
Import-Module AzureAD
Connect-AzureAD
Get-AzureADUser -ObjectId "CQPresse@precisionglobal.com" | Select-Object ImmutableId

--> No ImmutableId


*** Powershell 5.1 *** On your DC
$guid = (Get-ADUser -Identity "CPresse").ObjectGUID
$immutableID = [System.Convert]::ToBase64String($guid.ToByteArray())
$immutableID


*** Powershel 5.1 *** On my Computer
Set-AzureADUser -ObjectId "CQPresse@precisionglobal.com" -ImmutableId "m+iWHXVVfU6GDgq2+OkTyw=="