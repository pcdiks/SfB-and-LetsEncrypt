#Requires -Version 4
#Requires -RunAsAdministrator

<#
.SYNOPSIS
Request certificates from Let's Encrypt for use by Skype for Business/Lync Edge Server.

.DESCRIPTION
This scripts Will request three certificates from Let's Encrypt and assign them to the three Edge Server roles. This script does not request certificates for the XMPP role.

.PARAMETER PfxPassword
Password for protecting the requested certificate

.PARAMETER SipFQDN
Domain FQDN name(s) for the Access Edge External role

.PARAMETER WebFQDN
Domain FQDN name for the Data Edge External Role

.PARAMETER AvFQDN
Domain FQDN name for the Audio Video Authentication

.PARAMETER Live
If you add this parameter, you will get the 'real' certificates. Without this parameter you will get certificates from the test Let's Encrypt Server. Start this script first without the Live switch and if everything works as expected, add the Live switch to get the 'real' certificates.

.EXAMPLE
.\Update-Certificates.ps1 -PfxPassword V3ryStr0ngP@ssword -SipFQDN sip.mydomainname.net -WebFQDN web.mydomainname.net -AvFQDN av.mydomainname.net
In this example, we will request three certificates for mydomainname.net and protect the pfx files with the supplied password.

.INPUTS
None

.OUTPUTS
Certificates

.LINK
1: Mongoose Web Server: https://cesanta.com/binary.html
2: le64.exe: https://github.com/do-know/Crypt-LE/releases

.FUNCTIONALITY
Request certificates from Let's Encrypt to be used by the Skype for Business/Lync Edge Server
#>

[CmdletBinding()]
param
(
    # Target directory for exporting credentials and certificate.
    [parameter(Mandatory = $true,
        HelpMessage = "Enter a password to protect the pfx file")]
    [string]$PfxPassword,

    # Log directory for transcript logs, i.e. "\\server\share\logs"
    [Parameter(Mandatory = $true,
        HelpMessage = "Enter the (comma seperated) name(s) for the Access Edge External Role (sip.mydomainname.net)")]
    [string]$SipFQDN,

    # Include the service account name that will be used to execute commands or scripts
    [Parameter(Mandatory=$true,
        HelpMessage = "Enter the name for the Data Edge External Role (web.mydomainname.net)")]
    [string]$WebFQDN,

    [Parameter(Mandatory = $true,
        HelpMessage = "Enter the name for the Audio Video Authentication Role (av.mydomainname.net)")]
    [string]$AvFQDN,

    [Parameter(Mandatory = $false)]
    [switch]$Live
) # end param

#Check if prerequisites are available
$ScriptPath = $PSScriptRoot

if ( -not (Test-Path -Path (Join-Path -Path $ScriptPath -ChildPath "le64.exe"))){
    Write-Host "Download le64.exe and place it in folder: $ScriptPath" -ForegroundColor Red
    exit
} else {
    Write-Host "le64 executable found" -ForegroundColor Green
}

if ( -not (Test-Path -Path (Join-Path -Path $ScriptPath -ChildPath "mongoose-free*.exe"))){
    Write-Host "Download Mongoose Free and place it in the folder: $ScriptPath" -ForegroundColor Red
    exit
} else {
    Write-Host "mongoose executable found" -ForegroundColor Green
}

#Check for mongoose.conf and create one if not found
$ConfFile = Join-Path -Path $ScriptPath -ChildPath "mongoose.conf"
if ( -not (Test-Path -Path $ConfFile)) {
    Write-Host "Creating mongoose.conf file..." -ForegroundColor Yellow
    $WWWPath = Join-Path -Path $ScriptPath -ChildPath www
    $ConfContent = "document_root $WWWPath
    listening_port 80
    "
    $ConfContent | Out-File $ConfFile -Encoding ascii
}

#Check if www folders exist
if ( -not (Test-Path -Path (Join-Path -Path $ScriptPath -ChildPath "www\.well-known\acme-challenge"))) {
    New-Item -Name (Join-Path -Path $ScriptPath -ChildPath "www\.well-known\acme-challenge") -ItemType Directory
}

#Create and or enalbe Firewall Rules
if (Get-NetFirewallRule -DisplayName "Allow Mongoose" -ErrorAction SilentlyContinue ) {
    Enable-NetFirewallRule -DisplayName "Allow Mongoose"
} else {
    New-NetFirewallRule -DisplayName "Allow Mongoose" -Direction Inbound -Program "$ScriptPath\mongoose-free-6.9.exe" -Action Allow
}

Set-Location $ScriptPath

#Force Skype for Business Services to stop
Write-Host "Stopping Skype for Business services..." -ForegroundColor Green
Stop-CsWindowsService -NoWait -Force

Write-Host "Starting Web Server..." -ForegroundColor Green
Start-Process -FilePath $ScriptPath\mongoose-free-6.9.exe
Start-Sleep -Seconds 10
Write-Host "Start request for $SipFQDN" -ForegroundColor Green
if ($Live) {
    .\le64.exe --key account.key --csr sip.csr --csr-key sip.key --crt sip.crt --domains $SipFQDN --path $ScriptPath\www\.well-known\acme-challenge\ --generate-missing --unlink --export-pfx $PfxPassword --live
    certutil -f -p $PfxPassword -importpfx .\sip.pfx NoRoot
    if ($SipFQDN -match ",") {
        $Certificate = $SipFQDN.Split(",")[0]
        $ImportSip = Get-ChildItem Cert:\LocalMachine\my\ | Where-Object {$_.Subject -eq "cn=$SipFQDN[0]" -and $_.Issuer -eq "CN=Let's Encrypt Authority X3, O=Let's Encrypt, C=US" -and $_.NotBefore.ToString("yyyyMMdd") -eq (Get-Date -Format yyyyMMdd)}    
    } else {
        $ImportSip = Get-ChildItem Cert:\LocalMachine\my\ | Where-Object {$_.Subject -eq "cn=$SipFQDN" -and $_.Issuer -eq "CN=Let's Encrypt Authority X3, O=Let's Encrypt, C=US" -and $_.NotBefore.ToString("yyyyMMdd") -eq (Get-Date -Format yyyyMMdd)}
    }
    Set-CsCertificate -Thumbprint $ImportSip.Thumbprint -Type AccessEdgeExternal    
} else {
    .\le64.exe --key account.key --csr sip.csr --csr-key sip.key --crt sip.crt --domains $SipFQDN --path $ScriptPath\www\.well-known\acme-challenge\ --generate-missing --unlink --export-pfx $PfxPassword
    certutil -f -p $PfxPassword -importpfx .\sip.pfx NoRoot
}

Write-Host "Start request for $WebFQDN" -ForegroundColor Green
if ($Live){
    .\le64.exe --key account.key --csr web.csr --csr-key web.key --crt web.crt --domains $WebFQDN --path C:\StartReady\LetsEncrypt\www\.well-known\acme-challenge\ --generate-missing --unlink --export-pfx $PfxPassword --live
    certutil -f -p $PfxPassword -importpfx .\web.pfx NoRoot
    $ImportSip = Get-ChildItem Cert:\LocalMachine\my\ | Where-Object {$_.Subject -eq "cn=$WebFQDN" -and $_.Issuer -eq "CN=Let's Encrypt Authority X3, O=Let's Encrypt, C=US" -and $_.NotBefore.ToString("yyyyMMdd") -eq (Get-Date -Format yyyyMMdd)}
    Set-CsCertificate -Thumbprint $ImportWeb.Thumbprint -Type DataEdgeExternal    
} else {
    .\le64.exe --key account.key --csr web.csr --csr-key web.key --crt web.crt --domains $WebFQDN --path C:\StartReady\LetsEncrypt\www\.well-known\acme-challenge\ --generate-missing --unlink --export-pfx $PfxPassword
    certutil -f -p $PfxPassword -importpfx .\web.pfx NoRoot
}

Write-Host "Start request for $AvFQDN" -ForegroundColor Green
if ($Live) {
    .\le64.exe --key account.key --csr av.csr --csr-key av.key --crt av.crt --domains $AvFQDN --path C:\StartReady\LetsEncrypt\www\.well-known\acme-challenge\ --generate-missing --unlink --export-pfx $PfxPassword --live
    certutil -f -p $PfxPassword -importpfx .\av.pfx NoRoot
    $ImportSip = Get-ChildItem Cert:\LocalMachine\my\ | Where-Object {$_.Subject -eq "cn=$AvFQDN" -and $_.Issuer -eq "CN=Let's Encrypt Authority X3, O=Let's Encrypt, C=US" -and $_.NotBefore.ToString("yyyyMMdd") -eq (Get-Date -Format yyyyMMdd)}
    Set-CsCertificate -Thumbprint $ImportAV.Thumbprint -Type AudioVideoAuthentication    
} else {
    .\le64.exe --key account.key --csr av.csr --csr-key av.key --crt av.crt --domains $AvFQDN --path C:\StartReady\LetsEncrypt\www\.well-known\acme-challenge\ --generate-missing --unlink --export-pfx $PfxPassword
    certutil -f -p $PfxPassword -importpfx .\av.pfx NoRoot    
}

Write-Host "Close firewall..." -ForegroundColor Green
Disable-NetFirewallRule -DisplayName "Allow Mongoose"
Write-Host "Stopping web server..." -ForegroundColor Green
Get-Process -Name "mongoose*" | Stop-Process

#Remove Old Certificates
Write-Host "Cleaning up old expired certificates..." -ForegroundColor Green
Get-ChildItem -Path cert:\LocalMachine\My -ExpiringInDays 0 | Remove-Item

#Remove old files
Write-Host "Cleaning up work files..." -ForegroundColor Green
Get-ChildItem (Join-Path -Path $ScriptPath -ChildPath *.pfx) | Remove-Item

#Start the Skype for Business Services
Write-Host "Starting Skype for Business services..." -ForegroundColor Green
Start-CsWindowsService
