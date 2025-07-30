
Clear-Host

#        _____  .__                 __    __________                                               .___
#       /  _  \ |  |   ____________/  |_  \______   \_____    ______ ________  _  _____________  __| _/
#      /  /_\  \|  | _/ __ \_  __ \   __\  |     ___/\__  \  /  ___//  ___/\ \/ \/ /  _ \_  __ \/ __ | 
#     /    |    \  |_\  ___/|  | \/|  |    |    |     / __ \_\___ \ \___ \  \     (  <_> )  | \/ /_/ | 
#     \____|__  /____/\___  >__|   |__|    |____|    (____  /____  >____  >  \/\_/ \____/|__|  \____ | 
#             \/          \/                              \/     \/     \/                          \/ 
#                  - Réalised by KERLOC'H Jean-Philip @ HashTag -
#                            - PrecisionGlobal @ 2025 -

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$DateThreshold = 10

$SMTPServer = "smtp.mailjet.com"
$SMTPPort = 587
$SMTPSender = "IT Notifications <itnotification@precisionglobal.com>"
$SMTPEncoding = [System.Text.Encoding]::UTF8
$SMTPUsername = "eee07b566173a29571ac17093642966a"
$SMTPPassword = ConvertTo-SecureString "d90fd2e3d95834d8468ca7c30a9f6ceb" -AsPlainText -Force
$SMTPCredential = New-Object System.Management.Automation.PSCredential ($SMTPUsername, $SMTPPassword)

[boolean]$SendReportAdmin = $true
$SendReportAdminEmail = "pvsa.it@precisionglobal.com"

Function Send-MailMessageForUser {
    Param(
        [Parameter(Mandatory=$true)][string]$SendMailUserGivenName,
        [Parameter(Mandatory=$true)][string]$SendMailUserSurname,
        [Parameter(Mandatory=$true)][string]$SendMailUserEmail,
        [Parameter(Mandatory=$true)][string]$SendMailUserPrincipalName,
        [Parameter(Mandatory=$true)][string]$SendMailUserPasswordExpirationDate,
        [Parameter(Mandatory=$true)][string]$SendMailUserPasswordDaysToExpiry
    )

    $SendMailBody = @"
<div style='font-family: Arial, sans-serif; line-height: 1.6; color: #333; padding: 20px;'>
<div style='text-align: center; margin-bottom: 30px;'>
    <img src='https://storageaccountprecision.blob.core.windows.net/precisionlogo/Transp_Logo%20300x137.png' alt='Logo Precision Global' style='width:300px;height:auto;margin-bottom:20px;'>
</div>

<p><strong>********* English Version ********</strong></p>
<p>Dear $SendMailUserGivenName,</p>
<p>In <b>$SendMailUserPasswordDaysToExpiry days</b>, the password for your account <b>$SendMailUserPrincipalName</b> will expire.<br>
Please remember to change it before it reaches the expiration date (<b>expiration date: $SendMailUserPasswordExpirationDate</b>).</p>
<p>We remind you that if you do not change this password before it expires, your account will be <b>locked</b>, and you will no longer have access to Precision resources until an administrator unlocks your account.</p>
<p>The simplest way to change your password is to press <b>Ctrl + Alt + Del</b> on your computer from your Windows desktop, then select the <b>'Change a password'</b> option.<br>
If you are a remote user, please launch the VPN first before changing the password.</p>
<p>Passwords must meet the following minimum requirements:</p>
<ul>
    <li>Not contain the user’s account name or parts of the user’s full name that exceed two consecutive characters</li>
    <li>Be at least eight characters in length</li>
    <li>Contain characters from three of the following four categories:
        <ul>
            <li>Uppercase letters (A through Z)</li>
            <li>Lowercase letters (a through z)</li>
            <li>Numbers (0 through 9)</li>
            <li>Non-alphabetic characters (e.g., !, $, #, %)</li>
        </ul>
    </li>
</ul>
<p>Best Regards,<br>The IT Team</p>
</div>
"@

    $SendMailObject = "$SendMailUserGivenName $SendMailUserSurname : your password is expiring !"

    Send-MailMessage -Verbose -SmtpServer $SMTPServer -Encoding $SMTPEncoding `
        -From $SMTPSender -To $SendMailUserEmail `
        -Subject $SendMailObject `
        -Body $SendMailBody -BodyAsHtml -Port $SMTPPort -Credential $SMTPCredential -UseSsl
}

$DateToday = (Get-Date).ToFileTime()
$DateWithThreshold = (Get-Date).AddDays($DateThreshold).ToFileTime()

$UsersInfos = Get-ADUser -Filter { (Enabled -eq $True) -and (PasswordNeverExpires -eq $False)} `
    -Properties "DisplayName", "mail", "msDS-UserPasswordExpiryTimeComputed" |
    Select-Object "GivenName", "Surname", "mail", "UserPrincipalName", "msDS-UserPasswordExpiryTimeComputed"

Write-Host "Utilisateurs récupérés : $($UsersInfos.Count)"
$UsersNotifList = @()

foreach ($User in $UsersInfos) {
    if (($User."msDS-UserPasswordExpiryTimeComputed" -lt $DateWithThreshold) -and ($User."msDS-UserPasswordExpiryTimeComputed" -gt $DateToday)) {
        $UserPasswordExpirationDate = [datetime]::FromFileTime($User."msDS-UserPasswordExpiryTimeComputed")
        $UserPasswordDaysToExpiry = [int](($User."msDS-UserPasswordExpiryTimeComputed" - $DateToday) / 864000000000)

        $UserObj = New-Object PSObject
        $UserObj | Add-Member NoteProperty GivenName $User.GivenName
        $UserObj | Add-Member NoteProperty Surname $User.Surname
        $UserObj | Add-Member NoteProperty Email $User.mail
        $UserObj | Add-Member NoteProperty UserPrincipalName $User.UserPrincipalName
        $UserObj | Add-Member NoteProperty PasswordExpirationDate ($UserPasswordExpirationDate.ToString("dd/MM/yyyy"))
        $UserObj | Add-Member NoteProperty PasswordDaysToExpiry $UserPasswordDaysToExpiry

        $UsersNotifList += $UserObj

        Send-MailMessageForUser -SendMailUserGivenName $User.GivenName `
            -SendMailUserSurname $User.Surname `
            -SendMailUserEmail $User.mail `
            -SendMailUserPrincipalName $User.UserPrincipalName `
            -SendMailUserPasswordExpirationDate ($UserPasswordExpirationDate.ToString("d MMMM yyyy")) `
            -SendMailUserPasswordDaysToExpiry $UserPasswordDaysToExpiry
    }
}

Write-Host "Users to be notified : $($UsersNotifList.Count)"

if ($SendReportAdmin -and ($UsersNotifList.Count -ne 0)) {
    $SendMailAdminBody = $UsersNotifList |
        ConvertTo-HTML -PreContent "Hello,<br><p>Here is the list of Active Directory accounts whose passwords expire in less than $DateThreshold days.</p>" |
        Out-String | ForEach-Object {
            $_ -replace "<table>", "<table style='border: 1px solid;'>" `
               -replace "<th>", "<th style='border: 1px solid; padding: 5px; background-color:#014B83; color:#fff;'>" `
               -replace "<td>", "<td style='padding: 10px;'>"
        }

    Send-MailMessage -Verbose -SmtpServer $SMTPServer -Encoding $SMTPEncoding `
        -From $SMTPSender -To $SendReportAdminEmail `
        -Subject "Summary - AD Password Expiration - $(Get-Date -Format dd/MM/yyyy)" `
        -Body $SendMailAdminBody -BodyAsHtml -Port $SMTPPort -Credential $SMTPCredential -UseSsl
}

if ($SendReportAdmin -and ($UsersNotifList.Count -eq 0)) {
    $Body = @"
Hi,<br><br>
No passwords are scheduled to expire within the next $DateThreshold days.<br><br>
Best regards,<br>
Your IT Department
"@

    Send-MailMessage -Verbose -SmtpServer $SMTPServer -Encoding $SMTPEncoding `
        -From $SMTPSender -To $SendReportAdminEmail `
        -Subject "Summary - No expiring passwords - $(Get-Date -Format dd/MM/yyyy)" `
        -Body $Body -BodyAsHtml -Port $SMTPPort -Credential $SMTPCredential -UseSsl
}