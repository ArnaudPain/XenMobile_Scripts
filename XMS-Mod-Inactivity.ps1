# Script to Modify Inactivity Timer
# From default 7 days to a new value
# Created by Arnaud Pain
# October, 2017
# Version 1.0

<#
.SYNOPSIS
Modify Inqctivity Timer on Server Properties

.DESCRIPTION
This script will log into a XenMobile Server, list the current value, ask for the new and apply it.

#>

#region connect to XMS
# Bypass certificate verification to enable access with XMS IP Address 
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

# Connect to XMS server 
$host.ui.RawUI.ForegroundColor = "White"
$XMS = Read-Host -Prompt 'Please provide url of the XMS Server'

# Function XMS-Test to verify FQDN if script run from Internet
$DNSName = $XMS
Function XMS-Test
{
trap [System.Management.Automation.MethodInvocationException]{
    write-host "Warning: " -ForegroundColor Red; 
	write-host "Host does not exist
Please verify the address provided" -Foregroundcolor Yellow; $host.ui.RawUI.ForegroundColor = "white"; exit}

write-host ([System.Net.Dns]::GetHostAddresses($XMS)>$null)
$host.ui.RawUI.ForegroundColor = "Green"
write-host "	Host verification successful"
$host.ui.RawUI.ForegroundColor = "White"
write-host " "
}

# Function to check if port 4443 is opened
Function Port-Test
{
$test=(New-Object System.Net.Sockets.TcpClient).Connect($XMS, 4443) 
trap [System.Management.Automation.MethodInvocationException]{
    write-host "Warning: " -ForegroundColor Red; 
	write-host "Port 4443 is not opened" -Foregroundcolor Yellow; $host.ui.RawUI.ForegroundColor = "white"; exit}

$host.ui.RawUI.ForegroundColor = "Green"
write-host "	Port 4443 is open"
$host.ui.RawUI.ForegroundColor = "White"
}

# Check if XMS server can be resolved 
$host.ui.RawUI.ForegroundColor = "Yellow"
write-host "Verifying Host:" $XMS
$host.ui.RawUI.ForegroundColor = "white"
XMS-Test

# Check if port 4443 is opened
$host.ui.RawUI.ForegroundColor = "Yellow"
write-host "Verifying if port 4443 is open for" $XMS
write-host " "
$host.ui.RawUI.ForegroundColor = "white"
Port-Test

# Get Login credentials
write-host "Please provide username and password"
$Credential = get-credential $null

# Check Credentials before continue
$log = '{{"login":"{0}","password":"{1}"}}'
$cred = ($log -f $Credential.UserName, $Credential.GetNetworkCredential().Password)

$headers=@{"Content-Type" = "application/json"}
$Url = "https://${XMS}:4443/xenmobile/api/v1/authentication/login"
$json=Invoke-RestMethod -Uri $url -Body $cred -ContentType application/json -Headers $headers -Method POST -Verbose:$false
$headers.add("auth_token",$json.auth_token)
trap {"" ;continue}

if ($json -eq $null)
{
 $host.ui.RawUI.ForegroundColor = "Red"
 write-host "Authentication failed - please verify your username and password."
 $host.ui.RawUI.ForegroundColor = "white"
 exit #terminate the script.
}
else
{
 $host.ui.RawUI.ForegroundColor = "Green"
 write-host "	Successfully authenticated with XMS Server"
 $host.ui.RawUI.ForegroundColor = "White"
}
#endregion

#region Retrieve List of Delivery Groups
$srvprop=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/serverproperties" -Headers $headers -Method Get
$host.ui.RawUI.ForegroundColor = "Yellow"
$count=$srvprop.allewproperties.length
Write-host "$(Get-Date): Server Properties"
$host.ui.RawUI.ForegroundColor = "white"

ForEach($allewproperties in $srvprop)
{
for ($v=0;$v -lt $count; $v++)
{
if($allewproperties.allewproperties[$v].name -match "device.inactivity.days.threshold")
{
#write-output $allewproperties.allewproperties[$v].id
$name = $allewproperties.allewproperties[$v].name 
write-host "Actual Value" $allewproperties.allewproperties[$v].value
write-host ""
$dn = $allewproperties.allewproperties[$v].displayname
$desc = $allewproperties.allewproperties[$v].description
write-host "Default Value" $allewproperties.allewproperties[$v].defaultvalue
}
else{continue}
}
}

$newval = Read-Host -Prompt 'Please provide the new Value you want to have'

$srvp=
'
{
  "name": "' + $name +'", 
  "value": "' + $newval +'",
  "displayName": "' + $dn +'",
  "description": "' + $desc +'"
}
'
$headers=@{"Content-Type" = "application/json"}
$Url = "https://${XMS}:4443/xenmobile/api/v1/authentication/login"
$json=Invoke-RestMethod -Uri $url -Body $log -ContentType application/json -Headers $headers -Method POST -Verbose:$false
$headers.add("auth_token",$json.auth_token)
$srvprop=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/serverproperties" -Body -$srvp -Headers $headers -Method Post
$host.ui.RawUI.ForegroundColor = "Yellow"
write-host "$(Get-Date): Server Property Changed, a reboot is needed"
$host.ui.RawUI.ForegroundColor = "White"

