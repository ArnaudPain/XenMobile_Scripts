# Script to revoke Inactive Devices
# 
# Created by Arnaud Pain
# October, 2017
# Version 1.0

<#
.SYNOPSIS
List inactive devices adn revoke to free licenses

.DESCRIPTION
This script will log into a XenMobile Server, list the inactive devices, ask for confirmation to revoke, and revoke

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

#endregion

# Define function to revoke device
Function Revoke-Devices
{
 $host.ui.RawUI.ForegroundColor = "Yellow"
Write-host "$(Get-Date): Revoke Device ID"$id
$rev=
'
['+$id+']
'
$revoke=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/device/revoke"-Body $rev -Headers $headers -Method Post
$host.ui.RawUI.ForegroundColor = "Green"
Write-Host "Inactive Devices have been revoked"
$host.ui.RawUI.ForegroundColor = "white"
}

#region list current value for inactivy
$srvprop=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/serverproperties" -Headers $headers -Method Get
$host.ui.RawUI.ForegroundColor = "Yellow"
$count=$srvprop.allewproperties.length
Write-host "$(Get-Date): Get Server Property for Inactivity"
$host.ui.RawUI.ForegroundColor = "white"

ForEach($allewproperties in $srvprop)
{
for ($v=0;$v -lt $count; $v++)
{
if($allewproperties.allewproperties[$v].name -match "device.inactivity.days.threshold")
{
$id = $allewproperties.allewproperties[$v].id
$name = $allewproperties.allewproperties[$v].value
write-host "Actual Value	" $allewproperties.allewproperties[$v].value "Days"
}
else{continue}
}
}
#endregion

#region Retrieve List of Inactive Devices
$devBody=
'
{
"start": 0,
"limit": 10000
}
'
$dev=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/device/filter -Body $devBody -Headers $headers -Method Post -Verbose:$false

$count = $dev.matchedRecords
$host.ui.RawUI.ForegroundColor = "Yellow"
$inactive = 0
$inact = 0
Write-host "$(Get-Date): List Inactive Devices"
$host.ui.RawUI.ForegroundColor = "white"
for($v=0;$v -lt $count;$v++)
{
$test = $dev.filteredDevicesDataList[$v].inactivityDays
$inact++
if([int]$test -gt $name)
{
$inactive++
}
}

$answer = read-host "You have $inactive inactive Devices. Confirm the revocation of the $inactive inactive Devices (y/n)?"

if($answer -eq "y")
{
for($v=0;$v -lt $count;$v++)
{
$test = $dev.filteredDevicesDataList[$v].inactivityDays
if([int]$test -gt $name)
{
$id = $dev.filteredDevicesDataList[$v].id
Revoke-Devices
}
}
}
elseif
($answer -eq "n")
{
$host.ui.RawUI.ForegroundColor = "red"
write-host "You have chosen to not revoke the Devices"
$host.ui.RawUI.ForegroundColor = "white"
}
elseif
(!$answer)
{
$host.ui.RawUI.ForegroundColor = "red"
write-host "An anwser is needed!"
$host.ui.RawUI.ForegroundColor = "white"
}


