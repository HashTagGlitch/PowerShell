Clear-Host
#        _____                    __      __ .__                             ________    ________ .________
#       /  _  \    ____   ___.__./  \    /  \|  |__    ____  _______   ____  \_____  \  /  _____/ |   ____/
#      /  /_\  \  /    \ <   |  |\   \/\/   /|  |  \ _/ __ \ \_  __ \_/ __ \   _(__  < /   __  \  |____  \ 
#     /    |    \|   |  \ \___  | \        / |   Y  \\  ___/  |  | \/\  ___/  /       \\  |__\  \ /       \
#     \____|__  /|___|  / / ____|  \__/\  /  |___|  / \___  > |__|    \___  >/______  / \_____  //______  /
#             \/      \/  \/            \/        \/      \/              \/        \/        \/        \/ 
#
#                  - Réalisé par KERLOC'H Jean-Philip @ HashTag -
#                            - APST37 @ 2025 -


# ---------------------------------------------#
# A faire uniquement à la premiere execution   #
# ---------------------------------------------#

# Register-PnPEntraIDApp -ApplicationName "PnP-AnyWhere365" -Tenant apst37fr.onmicrosoft.com -Interactive
# Set-ExecutionPolicy Bypass -Scope CurrentUser -Force



# ------------------------------------------------#
# Vérification de l'état du repository PSGallery  #
# ------------------------------------------------#

$repository = Get-PSRepository -Name 'PSGallery' -ErrorAction SilentlyContinue

if ($repository -and $repository.InstallationPolicy -eq 'Trusted') {
    Write-Host "Le repository 'PSGallery' est déjà défini comme Trusted." -ForegroundColor Green
    Start-Sleep -Seconds 2
} else {
    Write-Host "Le repository 'PSGallery' n'est pas Trusted. Mise à jour en cours..." -ForegroundColor Yellow
    Start-Sleep -Seconds 2
    try {
        Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
        Write-Host "Le repository 'PSGallery' a été défini en mode Trusted avec succès." -ForegroundColor Green
        Start-Sleep -Seconds 2
    } catch {
        Write-Host "Erreur lors de la mise à jour du repository 'PSGallery'. Vérifiez vos permissions." -ForegroundColor Red
        Start-Sleep -Seconds 2
        exit
    }
}

# ------------------------------------------------#
# Vérification et installation de PnP.PowerShell  #
# ------------------------------------------------#

# Vérification si le module est installé
$moduleName = "PnP.PowerShell"
if (Get-Module -ListAvailable -Name $moduleName) {
    Write-Host "Le module '$moduleName' est déjà installé." -ForegroundColor Green
    Start-Sleep -Seconds 2
} else {
    Write-Host "Le module '$moduleName' n'est pas installé. Installation en cours..." -ForegroundColor Yellow
    Start-Sleep -Seconds 2
    try {
        Install-Module -Name $moduleName -Force -Confirm:$false
        Write-Host "Le module '$moduleName' a été installé avec succès." -ForegroundColor Green
        Start-Sleep -Seconds 2
    } catch {
        Write-Host "Erreur lors de l'installation du module '$moduleName'. Vérifiez vos permissions." -ForegroundColor Red
        Start-Sleep -Seconds 2
        exit
    }
}


# ---------------------------------------#
# Vérification de la version PowerShell  #
# ---------------------------------------#

# Récupérer la version actuelle de PowerShell
$psVersion = $PSVersionTable.PSVersion.Major

if ($psVersion -lt 7) {
    Write-Host "PowerShell 7 est requis. Installation en cours..." -ForegroundColor Yellow
    # Vérifier si winget est installé
    if (Get-Command winget -ErrorAction SilentlyContinue) {
        # Installer PowerShell 7 avec winget
        winget install --id Microsoft.PowerShell --source winget --accept-source-agreements --accept-package-agreements
        Write-Host "Installation terminée. Veuillez redémarrer votre terminal et exécuter à nouveau le script." -ForegroundColor Green
        Start-Sleep -Seconds 2
    } else {
        Write-Host "Winget n'est pas installé sur ce système. Veuillez l'installer manuellement depuis le site Microsoft." -ForegroundColor Red
        Start-Sleep -Seconds 2
    }
    exit
}

Write-Host "PowerShell 7 détecté. Poursuite de l'exécution du script..." -ForegroundColor Green

# ----------------------------#
# Connexion à SharePoint      #
# ----------------------------#

# Récupération dynamique du répertoire utilisateur
$userProfile = $env:USERPROFILE

# Définition du chemin du certificat en fonction de l'utilisateur connecté
$secureCertificatePass = "$userProfile\OneDrive - APST37-SSTI\Documents\PnP-SharePoint\PnP-AnyWhere365.pfx"

# Vérification si le fichier de certificat existe
if (-Not (Test-Path $secureCertificatePass)) {
    Write-Host "Le fichier de certificat n'existe pas : $secureCertificatePass" -ForegroundColor Red
    Start-Sleep -Seconds 2
    exit
}
Write-Host "Chemin du certificat utilisé : $secureCertificatePass" -ForegroundColor Green
Start-Sleep -Seconds 2

# Chemin du certificat PFX
$siteUrl = "https://apst37fr.sharepoint.com/sites/ContactCenterFlex365/ucc_production"
$clientId = "7e11a2cd-eb9f-4aff-a289-4af7282774a4"
$tenant = "apst37.fr"

try {
    Connect-PnPOnline -Url $siteUrl -ClientId $clientId -CertificatePath $secureCertificatePass -Tenant $tenant
    Write-Host "Connexion à SharePoint réussie." -ForegroundColor Green
    Start-Sleep -Seconds 2
} catch {
    Write-Host "Échec de la connexion à SharePoint. Vérifiez les informations d'identification." -ForegroundColor Red
    Start-Sleep -Seconds 2
    exit
}

# ----------------------------#
# Fonction pour récupérer la prochaine valeur unique wsp_ucc_Order #
# ----------------------------#

function Get-NextOrderValue {
    param (
        [string]$listName,
        [int]$startOrder = 100
    )
    
    try {
        $items = Get-PnPListItem -List $listName -Fields "wsp_ucc_Order"
        $usedOrders = @($items | ForEach-Object { $_.FieldValues["wsp_ucc_Order"] }) | Sort-Object
        $nextOrder = $startOrder

        while ($usedOrders -contains $nextOrder) {
            $nextOrder++
        }
        return $nextOrder
    } catch {
        Write-Host "Erreur lors de la récupération des valeurs wsp_ucc_Order" -ForegroundColor Red
        exit
    }
}

# ----------------------------#
# Recherche de l'agent         #
# ----------------------------#
$listName = "Agents"
$emailToFind = Read-Host "Veuillez entrer l'adresse e-mail"
$item = Get-PnPListItem -List $listName -Fields "ID", "wsp_ucc_agent_teamsupn" | Where-Object { $_.FieldValues["wsp_ucc_agent_teamsupn"] -eq $emailToFind }

if ($item) {
    $AgentID = $item.FieldValues["ID"]
    Write-Host "L'adresse e-mail '$emailToFind' existe déjà avec l'ID: $AgentID" -ForegroundColor Green
} else {
    Write-Host "L'adresse e-mail '$emailToFind' n'existe pas. Création en cours..." -ForegroundColor Yellow

    # Récupération des informations utilisateur
    $telephone = Read-Host "Veuillez entrer le num de tel sous la forme +33xxx"
    $Centre = Read-Host "Veuillez entrer le Centre"
    $Nom = Read-Host "Veuillez entrer le Nom de Famille"
    $Prenom = Read-Host "Veuillez entrer le Prénom"

    # Obtenir la prochaine valeur unique pour wsp_ucc_Order
    $nextOrderValue = Get-NextOrderValue -listName $listName

    # Définition des valeurs pour l'ajout de l'agent
    $Values = @{
        "wsp_ucc_Order"              = $nextOrderValue
        "wsp_ucc_Formal"             = "False"
        "wsp_ucc_Agent"              = "sip:$emailToFind"
        "wsp_ucc_agent_teamsphone"   = "tel:$telephone"
        "wsp_ucc_agent_teamsupn"     = "$emailToFind"
        "_x0063_bt3"                 = "$Centre"
        "zfmk"                       = "$Nom"
        "pfrm"                       = "$Prenom"
    }

    # Ajout du nouvel agent
    try {
        $newItem = Add-PnPListItem -List $listName -Values $Values -ErrorAction Stop
        $AgentID = $newItem["ID"]
        Write-Host "L'utilisateur $emailToFind a été ajouté avec wsp_ucc_Order : $nextOrderValue" -ForegroundColor Green
    } catch {
        Write-Host "Erreur lors de l'ajout de l'utilisateur dans la liste SharePoint." -ForegroundColor Red
        exit
    }
}

# ----------------------------#
# Ajout des compétences        #
# ----------------------------#
$skillsListName = "SkillsPerAgent"
$skills = @(
    @{ skill = "13"; score = "100"; name = "AppelSortant" }
    @{ skill = "18"; score = "0"; name = "Intérim_Amboise" }
    @{ skill = "19"; score = "0"; name = "Intérim_Beaulieu" }
    @{ skill = "20"; score = "0"; name = "Intérim_Chinon" }
    @{ skill = "12"; score = "0"; name = "Intérim_Ecoparc" }
    @{ skill = "5"; score = "0"; name = "Secreatariat_Aeronef" }
    @{ skill = "1"; score = "0"; name = "Secreatariat_Amboise" }
    @{ skill = "2"; score = "0"; name = "Secreatariat_Beaulieu" }
    @{ skill = "6"; score = "0"; name = "Secreatariat_BlaisePascal" }
    @{ skill = "4"; score = "0"; name = "Secreatariat_Chinon" }
    @{ skill = "10"; score = "0"; name = "Secreatariat_EcoparcJaune" }
    @{ skill = "11"; score = "0"; name = "Secreatariat_EcoparcVert" }
)

foreach ($skill in $skills) {
    $Values = @{
        "wsp_ucc_spa_skill"  = $skill.skill
        "wsp_ucc_Score"      = $skill.score
        "wsp_ucc_spa_Agent"  = $AgentID
    }

    Add-PnPListItem -List $skillsListName -Values $Values -ErrorAction Stop
    Write-Host "Compétence ajoutée : $($skill.name) avec score $($skill.score)" -ForegroundColor Cyan
}

Write-Host "Toutes les compétences ont été ajoutées avec succès pour l'agent ID: $AgentID" -ForegroundColor Green

# ----------------------------------#
# Ajout dans Microsoft Teams        #
# ----------------------------------#

Connect-MicrosoftTeams
Install-Module -Name MicrosoftTeams -Force -AllowClobber

Set-CsPhoneNumberAssignment -Identity $emailToFind -PhoneNumber $telephone -PhoneNumberType DirectRouting
Set-CsPhoneNumberAssignment -Identity $emailToFind -EnterpriseVoiceEnabled $true
Grant-CsOnlineVoiceRoutingPolicy -Identity $emailToFind -PolicyName VRP-NoRestriction
Grant-CsTeamsCallingPolicy -Identity $emailToFind -PolicyName AllowCalling
Set-CsOnlineVoicemailUserSettings -Identity $emailToFind -VoicemailEnabled $false
Grant-CsTeamsAppSetupPolicy -Identity $emailToFind -PolicyName AnyWhere365


# Déconnexion
Disconnect-PnPOnline
Disconnect-MicrosoftTeams
Write-Host "Déconnexion réussie de SharePoint" -ForegroundColor Yellow
Write-Host "Déconnexion réussie de Microsoft Teams" -ForegroundColor Yellow