Connect-ExchangeOnline -UserPrincipalName jean-philip.kerloch@precisionglobal.com
Import-Module ExchangeOnlineManagement
Connect-IPPSSession
New-ComplianceSearch -Name "PhishingCleanup" -ExchangeLocation "All" -ContentMatchQuery 'from:"christy.greer@precisionglobal.com" AND subject:"Main Statement"'
Start-ComplianceSearch -Identity "PhishingCleanup"
Get-ComplianceSearch -Identity "PhishingCleanup"
Get-ComplianceSearch -Identity "PhishingCleanup" | Format-List -Property Items

New-ComplianceSearchAction -SearchName "PhishingCleanup" -Preview
Get-ComplianceSearchAction  -Identity "PhishingCleanup_Preview" | Format-List -Property Results

New-ComplianceSearchAction -SearchName "PhishingCleanup" -Purge -PurgeType HardDelete
Get-ComplianceSearchAction  -Identity "PhishingCleanup_Purge" | Format-List -Property Results

Remove-ComplianceSearch -Identity "PhishingCleanup"

#https://community.spiceworks.com/t/new-compliancesearchaction-preview-parameter-not-working/692900