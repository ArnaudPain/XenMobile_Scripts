#Requires -Version 3.0
#This File is in Unicode format.  Do not edit in an ASCII editor.

<#
.SYNOPSIS
    Determine the number of licenses used for Citrix XenMobile in the Cloud
	
.SAMPLE CODE
	Created with help from the original Citrix XenMobile Public API for REST Services	
	Downloadable at https://docs.citrix.com/content/dam/docs/en-us/xenmobile/10-4/Downloads/XenMobile-Public-API.pdf

.PARAMETER XMS
    XenMobile server FQDN name need port 4443 to be opened 
	
.PARAMETER Credential
	XenMobile local account with ADMIN role
	
	You are prompted for a password.

.NOTES
    Copyright (c) Arnaud Pain. All rights reserved.
#>

# Bypass certificate verification to enable access with XMS IP Address
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

# Login to XenMobile Server
$XMS = Read-Host -Prompt 'Please provide url of the XMS Server'

# Define Function XMS-Test
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

# Define Function to check if port 4443 is opened
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
$json=Invoke-RestMethod -Uri $url -Body $cred -Headers $headers -Method POST
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
 $host.ui.RawUI.ForegroundColor = "White"
 write-host "Successfully authenticated with XMS Server"
 write-host ""
 $host.ui.RawUI.ForegroundColor = "White"
}

# List number of Devices enrolled
$dev =
'
{ 
  "start": "0", 
  "limit": "100000"
}
'
$devices=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/device/filter" -Body $dev -Headers $headers -Method Post
# Declare an array to collect our result objects
$resultsarray =@()

# For every $contact held in the $contacts, do this loop
$new=0
$count = $devices.filteredDevicesDataList.length
$host.ui.RawUI.ForegroundColor = "Yellow"
write-host $count
$host.ui.RawUI.ForegroundColor = "White"
write-host "Device enrolled"

for ($v=0;$v -lt $count; $v++)
{
foreach($device in $devices.filteredDevicesDataList)
{
# Create a new custom object to hold our result.
$contactObject = new-object PSObject
# Add our data to $contactObject as attributes using the add-member commandlet
$contactObject | add-member -membertype NoteProperty -name "Enrolled User" -Value $devices.filteredDevicesDataList[$new].username
}
$new++
# Save the current $contactObject by appending it to $resultsArray ( += means append a new element to ‘me’)
$resultsarray += $contactObject
}

$resultsarray| Export-csv "devices.csv" -delimiter ";"

$host.ui.RawUI.ForegroundColor = "white"
$result = (import-csv devices.csv | Group -Property "Enrolled User").count
$host.ui.RawUI.ForegroundColor = "Yellow"
write-host $result
$host.ui.RawUI.ForegroundColor = "White"
write-host "Licenses used"
del devices.csv