Clear-Host
#
#
#            _________.__                        __________      .__        __ ________   ________.________       __________              
#           /   _____/|  |__ _____ _______   ____\______   \____ |__| _____/  |\_____  \ /  _____/|   ____/       \______   \____ ______  
#           \_____  \ |  |  \\__  \\_  __ \_/ __ \|     ___/  _ \|  |/    \   __\_(__  </   __  \ |____  \         |     ___/    \\____ \ 
#           /        \|   Y  \/ __ \|  | \/\  ___/|    |  (  <_> )  |   |  \  | /       \  |__\  \/       \        |    |  |   |  \  |_> >
#          /_______  /|___|  (____  /__|    \___  >____|   \____/|__|___|  /__|/______  /\_____  /______  /        |____|  |___|  /   __/ 
#                  \/      \/     \/            \/                       \/           \/       \/       \/                      \/|__|    
# 
#
#
#                  - Réalisé par KERLOC'H Jean-Philip @ HashTag -
#                            - PRECISION @ 2025 -


# ---------------------------------------------#
# A faire uniquement à la premiere execution   #
# ---------------------------------------------#
# Register-PnPEntraIDApp -ApplicationName "PnP-Microsoft365" -Tenant "precisionglobalmigration.onmicrosoft.com"
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
$secureCertificatePass = "$userProfile\OneDrive - Precision Global\Documents\PnP-SharePoint\PnP-Microsoft365.pfx"

# Vérification si le fichier de certificat existe
if (-Not (Test-Path $secureCertificatePass)) {
    Write-Host "Le fichier de certificat n'existe pas : $secureCertificatePass" -ForegroundColor Red
    Start-Sleep -Seconds 2
    exit
}
Write-Host "Chemin du certificat utilisé : $secureCertificatePass" -ForegroundColor Green
Start-Sleep -Seconds 2

# Chemin du certificat PFX
$siteUrl = "https://precisionglobalmigration.sharepoint.com"
$clientId = "5ebb53ec-b0a0-45cf-b9d2-a8f9db1d62fd"
$tenant = "precisionglobalmigration.onmicrosoft.com"
$Thumbprint ="5C26D932379DBBE4E459A57926C8147A6CE1AACD"

try {
    #Connect-PnPOnline -Tenant $tenant -Url $siteUrl -ClientId $clientId -Thumbprint $Thumbprint
    Connect-PnPOnline -Url $siteUrl -ClientId $clientId -CertificatePath $secureCertificatePass -Tenant $tenant
    Write-Host "Connexion à SharePoint réussie." -ForegroundColor Green
    Start-Sleep -Seconds 2
} catch {
    Write-Host "Échec de la connexion à SharePoint. Vérifiez les informations d'identification." -ForegroundColor Red
    Start-Sleep -Seconds 2
    exit
}




####################### REMOVE Sites Corbeille ############################

Get-PnPTenantDeletedSite | ForEach-Object {
    Write-Host "Suppression du site supprimé : $($_.Url)" -ForegroundColor Yellow
    Remove-PnPTenantDeletedSite -Url $_.Url -Force
}

##########################################################################