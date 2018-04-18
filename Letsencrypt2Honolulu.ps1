$SSHHost = "ssh.host.name"
$SSHKey = "ssh.key.pem"
$SSHKeyPassword = "ssh-key-secret"
$SSHUser = "ssh-user"
$CertificateCN = "certificate-common-name-to-use"
$ACMEPath = "/path/to/acme-v01.api.letsencrypt.org/sites"
$WindowsAdminCenterPort = "444"
$WindowsAdminCenterRandomGUID = "{38700e7c-87eb-471b-ab40-07f221d35b43}"

$SSHKeyPasswordSecString = ConvertTo-SecureString "$SSHKeyPassword" -AsPlainText -Force
$SSHCredentials = New-Object System.Management.Automation.PSCredential ("$SSHUser", $SSHKeyPasswordSecString)



Write-Host -ForegroundColor Cyan "[1/5] Download crt+key"
Get-SCPFolder -ComputerName $SSHHost -KeyFile "$SSHKey" -Credential $SSHCredentials -AcceptKey -RemoteFolder "$CertificateCN" -LocalFolder "$PSScriptRoot"

Write-Host -ForegroundColor Cyan "[2/5] Convert crt+key to PFX"
$env:Home = $PSScriptRoot
& .\openssl.exe pkcs12 -export -out "$CertificateCN.pfx" -inkey "$CertificateCN.key" -in "$CertificateCN.crt" -passout pass:

Write-Host -ForegroundColor Cyan "[3/5] Import PFX and remove temporary certificate files"
Import-PfxCertificate "$ACMEPath/$CertificateCN.pfx" -CertStoreLocation "Cert:\LocalMachine\My"
Get-ChildItem -Path "$PSScriptRoot" -Filter "$CertificateCN" | Remove-Item -Force

Write-Host -ForegroundColor Cyan "[4/5] Set IIS certificate"
$CertificateThumbprint = (Get-PfxCertificate -FilePath "$CertificateCN.pfx").Thumbprint
$binding = Get-WebBinding -Protocol "https"
$binding.AddSslCertificate($CertificateThumbprint, "my")

Write-Host -ForegroundColor Cyan "[4/5] Set Windows Admin Center certificate"
Stop-Service ServerManagementGateway
netsh http delete sslcert ipport=0.0.0.0:$WindowsAdminCenterPort
netsh http add sslcert ipport=0.0.0.0:$WindowsAdminCenterPort certhash=$CertificateThumbprint appid= '$WindowsAdminCenterRandomGUID'
Start-Service ServerManagementGateway