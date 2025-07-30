## création du certificat
# === PARAMÈTRES PERSONNALISABLES ===
$certName = "PnP-Microsoft365"
$pfxPath = "C:\Certs\PnP-Microsoft365.pfx"
$cerPath = "C:\Certs\PnP-Microsoft365.cer"
$certPasswordPlain = "4gurg7Xj8bvYEQ5VQT932"  # À modifier si besoin
$certPassword = ConvertTo-SecureString $certPasswordPlain -AsPlainText -Force

# === PRÉPARATION DU DOSSIER ===
New-Item -ItemType Directory -Path "C:\Certs" -Force | Out-Null

# === CRÉATION DU CERTIFICAT AUTO-SIGNÉ ===
$cert = New-SelfSignedCertificate `
    -Subject "CN=$certName" `
    -KeySpec Signature `
    -KeyExportPolicy Exportable `
    -KeyLength 2048 `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -NotAfter (Get-Date).AddYears(25) `
    -FriendlyName $certName

# === EXPORT EN PFX (avec mot de passe) ===
Export-PfxCertificate `
    -Cert "Cert:\CurrentUser\My\$($cert.Thumbprint)" `
    -FilePath $pfxPath `
    -Password $certPassword

# === EXPORT EN .CER (pour Azure App) ===
Export-Certificate `
    -Cert "Cert:\CurrentUser\My\$($cert.Thumbprint)" `
    -FilePath $cerPath


$certPassword = ConvertTo-SecureString "4gurg7Xj8bvYEQ5VQT932
" -AsPlainText -Force