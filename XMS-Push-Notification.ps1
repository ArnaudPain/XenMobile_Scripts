# Script to send Notification message to Delivery Group in XenMobile
# Created by Arnaud Pain
# September, 2017
# Version 1.2

<#
.SYNOPSIS
Send Notification to Delivery Group in XenMobile

.DESCRIPTION
This script will log into a XenMobile Server, list the Delivery Groups, ask for: DG, Message and Sound

Add September 29, 2017
Validate credential before continue, you must use account created on initial configuration in CLI for Web-based administration interface
Or a local account with role ADMIN

Add October 5, 2017
List Delivery group is limited by default to 10, change made to ensure to list/view all configured Delivery Groups (up to 1000).
#>

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

# Retrieve List of Delivery Groups
$dg=
'
{ 
  "start": "0", 
  "limit": "1000"
}
'
$dgroup=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/deliverygroups/filter" -Body $dg -Headers $headers -Method Post
$host.ui.RawUI.ForegroundColor = "Yellow"
Write-host "$(Get-Date): Extract Delivery Groups"
$host.ui.RawUI.ForegroundColor = "white"
write-host "Available Delivery Groups:"
$host.ui.RawUI.ForegroundColor = "green"
$new=0
$count = $dgroup.dglistdata.dglist.length
for ($v=0;$v -lt $count; $v++)
{
foreach($dglistdata in $dgroup)
{
write-host $dglistdata.dglistdata.dglist[$new].name
$new++
}
}

$host.ui.RawUI.ForegroundColor = "white"
$dgroup = Read-Host -Prompt "Please provide Delivery Group Name for which notification will be sent"
$message = Read-Host -Prompt "Please provide the message to be sent"
$sound = Read-Host -Prompt "Please provide the sound to be played (Congos,Sheep,Casino,AlarmHorn,Alert_Tone: => Case Sensitive)"

# Function Send Notification
Function Send-notification-iOS
{
$notification =
'
{ 
	"to": [ 
	{ "deviceId": "' + $device.id +'",
	  "osFamily": "iOS" 
	}
    ], 
	"agentMessage": "' +$message +'", 
	"smtp": "false", 
	"sms": "false", 
	"agent": "true",
	"templateId": "-1", 
	"agentCustomProps": { 
		"sound": "' + $sound +'.wav" 
	} 
}
'
$Notif=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/device/notify" -Body $notification -Headers $headers -Method Post
}

Function Send-notification-Android
{
$notification =
'
{ 
	"to": [ 
	{ "deviceId": "' + $device.id +'",
	  "osFamily": "ANDROID" 
	}
    ], 
	"agentMessage": "' +$message +'", 
	"smtp": "false", 
	"sms": "false", 
	"agent": "true",
	"templateId": "-1", 
	"agentCustomProps": { 
		"sound": "' + $sound +'.wav" 
	} 
}
'
$Notif=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/device/notify" -Body $notification -Headers $headers -Method Post
}

# Retrieve list of Devices
$devices=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/device/filter" -Body '{}' -Headers $headers -Method Post

foreach($device in $devices.filteredDevicesDataList)
{
$url = "https://${XMS}:4443/xenmobile/api/v1/device/" + $device.id + "/deliverygroups"
$dg=Invoke-RestMethod -Uri $url -Headers $headers -Method Get
foreach($deliveryGroups in $dg)
{
If($dg.deliveryGroups.name -match $dgroup)
{
$new=0
If($devices.filteredDevicesDataList[$new].platform -eq "iOS")
{
Send-notification-iOS
$new++
}
Send-notification-Android
}
}
}

