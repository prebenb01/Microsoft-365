$pwd = ConvertTo-SecureString "REDACTED" -AsPlainText -Force

New-PnPAzureCertificate `
    -CommonName "Cert" `
    -ValidYears 3 `
    -CertificatePassword $pwd `
    -OutPfx ".\cert.pfx" `
    -OutCert ".\cert.cer"

# Base64 the pfx for the app setting — single line, no wraps
[Convert]::ToBase64String([IO.File]::ReadAllBytes(".\cert.pfx")) | Out-File ".\cert-pfx-base64.txt"