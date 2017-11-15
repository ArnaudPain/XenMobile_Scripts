# Script to Add Actions to a Delivery Group
# Created by Arnaud Pain
# September, 2017
# Version 1.0

<#
.SYNOPSIS
Add Actions to Delivery Group in XenMobile

.DESCRIPTION
This script will log into a XenMobile Server, list the Delivery Groups, ask for: DG, nb of Actions to add and name of each

October 24, 2017 Modify script to keep track of current configuration to not delete it during the Add-Action

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

#region List Actions
$host.ui.RawUI.ForegroundColor = "white"
$dgname = Read-Host -Prompt "Please provide the Name of the Delivery Group"

Write-host "Actions already configured for"$dgname":"
$host.ui.RawUI.ForegroundColor = "Green"

$new=0
$count = $dgroup.dglistdata.dglist.length
foreach($dglistdata in $dgroup)
{
for ($v=0;$v -lt $count; $v++)
{
if ($dglistdata.dglistdata.dglist[$new].name -eq $dgname)
{
write-output $dglistdata.dglistdata.dglist[$new].smartActions.name
write-output $dglistdata.dglistdata.dglist[$new].smartActions.name >> currentact.txt
}
$new++
}
}
#endregion

$new=0
$count = $dgroup.dglistdata.dglist.length
for ($v=0;$v -lt $count; $v++)
{
foreach($dglistdata in $dgroup)
{
write-output $dglistdata.dglistdata.dglist[$new].name >>dglist.txt
$new++
}
}
#endregion


#endregion

#region List Applications configure for the Delivery Group
$new=0
$count = $dgroup.dglistdata.dglist.length
foreach($dglistdata in $dgroup)
{
for ($v=0;$v -lt $count; $v++)
{
if ($dglistdata.dglistdata.dglist[$new].name -eq $dgname)
{
write-output $dglistdata.dglistdata.dglist[$new].applications.name >> currentapp.txt
write-output $dglistdata.dglistdata.dglist[$new].applications.required >> currentappreq.txt
}
$new++
}
}

$file2 = Get-content currentapp.txt
$file1 = Get-content currentappreq.txt
$cont = 0
foreach ($value in $file2)
{
$compname = $file1[$cont]
$newname = "$value," + "$compname"
$newname >> result.txt
$cont = $cont+1
}

$file = get-content "result.txt"
$separator = ","
$parts = $file.split($separator)
$count = @(get-content "result.txt").Length
for ($v=0; $v -lt $count; $v++)
{
    $file = (get-content "result.txt")[$v]
    $separator = ","
    $parts = $file.split($separator)
    $z=
	'{
	"name": "'+ $parts[0] +'",
	"required": "'+ $parts[1] +'"
	},'
	$z>>currentresult.txt
}

# Modify file to delete last character ',' on last line
$app = get-content currentresult.txt
$app[0] = $app[0] -replace ','
$app[$app.length - 1 ] = $app[$app.length - 1 ] -replace ','
$app | set-content addapp.txt
#endregion


#region request information
$host.ui.RawUI.ForegroundColor = "white"
$count = Read-Host -Prompt "How many Actions do you want to add"

# Add each new Action to currentact.txt
for($v=0;$v -lt $count;$v++)
{
$w=$v+1
$name = Read-Host -Prompt "Please provide Actions $w name"
$name >> currentact.txt
}


# Create result.txt file with formated query
$acount=0
foreach ($value in get-content currentact.txt)
{
$z=
'{
"name": "'+ $value +'"
},'
$z>>result.txt
$acount++
}

# Modify file to delete last character ',' on last line
$pol = get-content result.txt
$pol[0] = $pol[0] -replace ','
$pol[$pol.length - 1 ] = $pol[$pol.length - 1 ] -replace ','
$pol | set-content addact.txt
#endregion

Function Disable-AllUsers
{
$headers=@{"Content-Type" = "application/json"}
$json=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/authentication/login -Body $Cred -Headers $headers -Method POST
$headers.add("auth_token",$json.auth_token)
$updAU = Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/deliverygroups/AllUsers/disable -Body '{}' -Headers $headers -Method Put
}

$host.ui.RawUI.ForegroundColor = "yellow"
write-host "$(Get-Date):  Update/Create Delivery Groups configuration"

#region retrieve Device Policies configured for the Delivery Group

	$new=0
	$count = $dgroup.dglistdata.dglist.length
	foreach($dglistdata in $dgroup)
		{
		for ($v=0;$v -lt $count; $v++)
			{
			if ($dglistdata.dglistdata.dglist[$new].name -eq $dgname)
				{
				write-output $dglistdata.dglistdata.dglist[$new].devicepolicies.name >> currentpol.txt
				}
			$new++
			}
		}
$file = get-content "currentpol.txt"
$separator = ","
$parts = $file.split($separator)
$count = @(get-content "currentpol.txt").Length
if($count -gt 1)
{
for ($v=0; $v -lt $count; $v++)
{
    $file = (get-content "currentpol.txt")[$v]
    $separator = ","
    $parts = $file.split($separator) 
    $z=
	'{
	"name": "'+ $parts[0] +'"
	},'
	$z>>currentpolresult.txt
}
}
Else
{
for ($v=0; $v -lt $count; $v++)
{
    $file = (get-content "currentpol.txt")[$v]
    $separator = ","
    $z=
	'{
	"name": "'+ $parts[0] +'"
	},'
	$z>>currentpolresult.txt
}
}

# Modify file to delete last character ',' on last line
$app = get-content currentpolresult.txt
$app[0] = $app[0] -replace ','
$app[$app.length - 1 ] = $app[$app.length - 1 ] -replace ','
$app | set-content addpol.txt
#endregion

#region retrieve Actions configured for the Delivery Group

	$new=0
	$count = $dgroup.dglistdata.dglist.length
	foreach($dglistdata in $dgroup)
		{
		for ($v=0;$v -lt $count; $v++)
			{
			if ($dglistdata.dglistdata.dglist[$new].name -eq $dgname)
				{
				write-output $dglistdata.dglistdata.dglist[$new].smartactions.name >> currentact.txt
				}
			$new++
			}
		}

$file = get-content "currentact.txt"
$separator = ","
$parts = $file.split($separator)
$count = @(get-content "currentact.txt").Length
if($count -gt 1)
{
for ($v=0; $v -lt $count; $v++)
{
    $file = (get-content "currentact.txt")[$v]
    $separator = ","
    $parts = $file.split($separator) 
    $z=
	'{
	"name": "'+ $parts[0] +'"
	},'
	$z>>currentactresult.txt
}
}
Else
{
for ($v=0; $v -lt $count; $v++)
{
    $file = (get-content "currentact.txt")[$v]
    $separator = ","
    $z=
	'{
	"name": "'+ $parts[0] +'"
	},'
	$z>>currentactresult.txt
}
}

# Modify file to delete last character ',' on last line
$app = get-content currentactresult.txt
$app[0] = $app[0] -replace ','
$app[$app.length - 1 ] = $app[$app.length - 1 ] -replace ','
$app | set-content addact.txt
#endregion				  
$pol = get-content addact.txt

#region retrieve Groups configured for the Delivery Group
$dgroup1=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/deliverygroups/$dgname -Headers $headers -Method Get
write-output $dgroup1.role.groups.uniqueName >> groupsun.txt
write-output $dgroup1.role.groups.UniqueID >> groupsui.txt
write-output $dgroup1.role.groups.domainName >> domainname.txt

$test = Get-content groupsun.txt
if($test)
{
$file1 = Get-content groupsun.txt
$file2 = Get-content groupsui.txt
$file3 = get-content domainname.txt
$cont = 0
foreach ($value in $file1)
{
$2 = $file2
$3 = $file3
$newname = "$value," + "$2," + "$3"
$newname >> groupsresult.txt
$cont = $cont+1
}

$file = get-content "groupsresult.txt"
$separator = ","
$parts = $file.split($separator)
$count = @(get-content "groupsresult.txt").Length
for ($v=0; $v -lt $count; $v++)
{
    $file = (get-content "groupsresult.txt")[$v]
    $separator = ","
    $z=
	'{
	"uniqueName": "'+ $parts[0] +'",
	"uniqueId": "'+ $parts[1] +'",
	"domainName": "'+ $parts[2] +'"
	},'
	$z>>groupe.txt
}

# Modify file to delete last character ',' on last line
$app = get-content groupe.txt
$app[0] = $app[0] -replace ','
$app[$app.length - 1 ] = $app[$app.length - 1 ] -replace ','
$app | set-content addgrp.txt
}
#endregion

#region Add Applications to Delivery Group
$test =test-path addgrp.txt
if($test -match "False")
{
$addpol = get-content addpol.txt
$addapp = get-content addapp.txt
$addact = get-content addact.txt
$devpol =
'
{ 
	"name": "' + $dgname +'",
	"applications": [
	'+ $addapp +'
	],
	"devicePolicies": [
	'+ $addpol +'
	],
	"smartActions" : [
	'+ $addact +'
	]
}
'
																																   

}
Else
{
$addpol = get-content addpol.txt
$addapp = get-content addapp.txt
$addact = get-content addact.txt
$addgrp = get-content addgrp.txt 

$devpol =
'
{ 
	"name": "' + $dgname +'",
	"applications": [
	'+ $addapp +'
	],
	"devicePolicies": [
	'+ $addpol +'
	],
	"smartActions" : [
	'+ $addact +'
	],
	"groups" :[
	'+ $addgrp +'
	]
}
'
}

$addaction=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/deliverygroups" -Body $devpol -Headers $headers -Method Put
if($test -match "True")
{
 del addact.txt
 del result.txt
 del currentact.txt
}
del groupsui.txt 
del groupsun.txt 
del addgrp.txt 
del groupe.txt 
del currentactresult.txt 
del addpol.txt 
del currentpolresult.txt 
del currentpol.txt 
del addapp.txt 
del currentresult.txt 
del groupsresult.txt 
del currentappreq.txt 
del currentapp.txt 
del domainname.txt
#enregion

del dglist.txt
$host.ui.RawUI.ForegroundColor = "Green"
write-host "$(Get-Date):	Action added to Delivery Groups"
$host.ui.RawUI.ForegroundColor = "White"
#endregion				 
