$AADComputer = "AzureAD.pvcazure.com"
$session = New-PSSession -ComputerName $AADComputer -Credential PVCAZURE\aaddsglobal@precisionglobal.com
#Invoke-Command -Session $session -ScriptBlock {Import-Module -Name 'ADSync'}
Invoke-Command -Session $session -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta}
Remove-PSSession $session