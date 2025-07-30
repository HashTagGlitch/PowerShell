Clear-Host

#        _____  .__                 __           __________                                               .___    _____       .___       
#       /  _  \ |  |   ____________/  |_  ____   \______   \_____    ______ ________  _  _____________  __| _/   /  _  \    __| _/_____  
#      /  /_\  \|  | _/ __ \_  __ \   __\/ __ \   |     ___/\__  \  /  ___//  ___/\ \/ \/ /  _ \_  __ \/ __ |   /  /_\  \  / __ |/     \ 
#     /    |    \  |_\  ___/|  | \/|  | \  ___/   |    |     / __ \_\___ \ \___ \  \     (  <_> )  | \/ /_/ |  /    |    \/ /_/ |  Y Y  \
#     \____|__  /____/\___  >__|   |__|  \___  >  |____|    (____  /____  >____  >  \/\_/ \____/|__|  \____ |  \____|__  /\____ |__|_|  /
#             \/          \/                 \/                  \/     \/     \/                          \/          \/      \/     \/ 
#                  - Realise par KERLOC'H Jean-Philip @ HashTag -
#                            - PrecisionGlobal @ 2025 -

# Forcer TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

### Parametres SMTP
$SMTPServer = "smtp.mailjet.com"
$SMTPPort = 587
$SMTPSender = "IT Notifications <itnotification@precisionglobal.com>"
$SMTPEncoding = [System.Text.Encoding]::UTF8
$SMTPUsername = "eee07b566173a29571ac17093642966a"
$SMTPPassword = ConvertTo-SecureString "d90fd2e3d95834d8468ca7c30a9f6ceb" -AsPlainText -Force
$SMTPCredential = New-Object System.Management.Automation.PSCredential ($SMTPUsername, $SMTPPassword)

$DateThreshold = 15
$DateToday = (Get-Date).ToFileTime()
$DateWithThreshold = (Get-Date).AddDays($DateThreshold).ToFileTime()

$GroupName = "Admins du Domaine"
$ReportList = @()

# Recuperation et analyse des comptes
Get-ADGroupMember -Identity $GroupName -Recursive |
Where-Object { $_.objectClass -eq "user" } |
ForEach-Object {
    $user = Get-ADUser -Identity $_.SamAccountName -Properties "DisplayName", "UserPrincipalName", "msDS-UserPasswordExpiryTimeComputed", "PasswordNeverExpires", "Enabled"
    
    $expiryRaw = $user."msDS-UserPasswordExpiryTimeComputed"
    if (
        $user.Enabled -eq $true -and
        $user.PasswordNeverExpires -ne $true -and
        $expiryRaw -and $expiryRaw -gt $DateToday -and $expiryRaw -lt $DateWithThreshold
    ) {
        $expiryDate = [datetime]::FromFileTime($expiryRaw)
        $daysLeft = [int](($expiryRaw - $DateToday) / 864000000000)

        $ReportList += [PSCustomObject]@{
            DisplayName     = $user.DisplayName
            UserPrincipal   = $user.UserPrincipalName
            ExpiryDate      = $expiryDate.ToString("dd/MM/yyyy")
            DaysRemaining   = $daysLeft
        }
    }
}

# Envoi de l'email uniquement s'il y a des comptes a notifier
if ($ReportList.Count -gt 0) {
    $HtmlBody = $ReportList |
        ConvertTo-Html -PreContent "<p>Hello,<br><br>Here is the list of <b>$GroupName</b> accounts whose passwords expire in less than  $DateThreshold days :</p>" |
        Out-String | ForEach-Object {
            $_ -replace "<table>", "<table style='border: 1px solid; border-collapse: collapse;'>" `
               -replace "<th>", "<th style='border: 1px solid; padding: 5px; background-color:#014B83; color:#fff;'>" `
               -replace "<td>", "<td style='border: 1px solid; padding: 5px;'>"
        }

    Send-MailMessage -Verbose -SmtpServer $SMTPServer -Encoding $SMTPEncoding `
        -From $SMTPSender -To "pvsa.it@precisionglobal.com" `
        -Subject "Report - Admin Accounts Password Expiration $(Get-Date -Format dd/MM/yyyy)" `
        -Body $HtmlBody -BodyAsHtml -Port $SMTPPort -Credential $SMTPCredential -UseSsl
}
else {
    $Body = @"
Hello,<br><br>
No passwords are scheduled to expire in the next  $DateThreshold days for group members <b>$GroupName</b>.<br><br>
Best regards,<br>
The IT Teams
"@

    Send-MailMessage -Verbose -SmtpServer $SMTPServer -Encoding $SMTPEncoding `
        -From $SMTPSender -To "pvsa.it@precisionglobal.com" `
        -Subject "Report - No expiring passwords $(Get-Date -Format dd/MM/yyyy)" `
        -Body $Body -BodyAsHtml -Port $SMTPPort -Credential $SMTPCredential -UseSsl
}
