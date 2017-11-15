# Script to Update MDX Apps
# 
# Created by Arnaud Pain
# October, 2017
# Version 1.0

<#
.SYNOPSIS
List MDX Apps and their respective version
Check the version of the MDX Apps on Public Store (Apple and Google)
If version available on public Store is newer, source is downloaded and Apps MDX is updated on XMS

.DESCRIPTION
This script will log into a XenMobile Server, list the MDX Version, Check version on public store and update if needed

.LIMITATION
For now there is no option to have the name of ShareFile to be written when upload and default ShareFile is used
You will need to change it manually after upgrade of the MDX

.UPDATE
Update November,6 2017 to reflect the new MDX from Citrix
#>

#Declare function update MDX if public is newer

Function Update-MDx
{
Function ConvertTo-PlainText( [security.securestring]$secure ) {
 $marshal = [Runtime.InteropServices.Marshal]
 $marshal::PtrToStringAuto( $marshal::SecureStringToBSTR($secure) )
}
$BASEURI="https://" + $XMS + ":4443"
$token=$json.auth_token
$headers=@{"Auth_token"=$token
			"Content-type"="multipart/form-data; boundary=----XenmobileScripting----"
			"Accept"= "*/*"
			"Accept-Encoding"="gzip, deflate"
			"Accept-Language"="en-US,en;q=0.8\r\n"
}
$URI=$BASEURI + "/xenmobile/api/v1/application/mobile/mdx/" + $platform + "/" + $id
$fileBin=[IO.File]::ReadAllBytes($filename)
$enc=[System.Text.Encoding]::GetEncoding("ISO-8859-1")
$fileenc=$enc.GetString($fileBin)

#Create a correctly formed DeliveryGroups object...
$deliverygroup=""
if(!$deliveryGroups)
{$deliveryGroups=""}
foreach ( $group in $deliverygroups.split(",")){
	if ($deliverygroup.Length -eq 0){
		$deliverygroup='["' + $group + '"'
		}
	else
		{
		$deliverygroup=$deliverygroup + ',"' + $group + '"'
		}
}
$deliverygroup=$deliverygroup + "]"

#Create a correctly formed Category object...
$category=""
foreach ( $cat in $categories.split(",")){
	if ($category.Length -eq 0){
		$category='["' + $cat + '"'
		}
	else
		{
		$category=$category + ',"' + $cat + '"'
		}
}
$category=$category + "]"

if(!$description)
{
$description = "Uploaded by Rest API Script"
}

$body = '------XenmobileScripting----
Content-Disposition:form-data;name="appInfo"

{
"name": "' + $appName + '",
"description": "' + $description + '",
"category": ' + $category + ',
"deliveryGroups":' + $deliverygroup + ',
"deploymentSchedule": {
"enableDeployment": true,
"deploySchedule": "NOW",
"deployDate": "",
"deployScheduleCondition": "EVERYTIME",
"deployInBackground": false
}
}
'

$body=$body + '------XenmobileScripting----
Content-Disposition: form-data;name="uploadFile";filename="' + (Get-Item $filename).basename + (Get-Item $filename).Extension + '"
Content-Type: application/octet-stream

' + $fileenc + '
------XenmobileScripting------
'



$return=try{
	Invoke-RestMethod -Headers $headers -Method Post -Uri $URI -Body $body
	} catch { $_.Exception.Response }
	
	if($return.StatusCode -ne $null)
	{
		Write-Output "Error! Errormessage:" $return.Message "`r`n" $return.InnerException
	}
	else
	{
		$host.ui.RawUI.ForegroundColor = "Green"
		Write-Output "$appname for $platform has been successfully uploaded!"
		$host.ui.RawUI.ForegroundColor = "Yellow"
	}

}
#endregion

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
if(!$XMS)
{
write-host "Host was not provided" -ForegroundColor Red; $host.ui.RawUI.ForegroundColor = "white"; exit}

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

#region download secure apps
$host.ui.RawUI.ForegroundColor = "Yellow"
Write-host "$(Get-Date): Create a Temporary folder C:\Temp\Secure Apps"
md "C:\Temp\Secure Apps" > $null
cd "C:\Temp\Secure Apps"
Write-host "$(Get-Date): Download latest version of MDX Apps"
Write-host "$(Get-Date): Download iOS MDX"
$url = "http://arnaudpain.com/download/1245/"
$output = "ios.zip"
Invoke-WebRequest -Uri $url -OutFile $output
Write-host "$(Get-Date): Download Android MDX"
$url = "http://arnaudpain.com/download/1248/"
$output = "android.zip"
Invoke-WebRequest -Uri $url -OutFile $output
#endregion

#region extract zip
Write-host "$(Get-Date): Extract MDX Apps"

$BackUpPath = "C:\Temp\Secure Apps\ios.zip"
$Destination = "C:\Temp\Secure Apps\iOS"
Add-Type -assembly "system.io.compression.filesystem"
[io.compression.zipfile]::ExtractToDirectory($BackUpPath, $destination)
del ios.zip

$BackUpPath = "C:\Temp\Secure Apps\android.zip"
$Destination = "C:\Temp\Secure Apps\Android"
Add-Type -assembly "system.io.compression.filesystem"
[io.compression.zipfile]::ExtractToDirectory($BackUpPath, $destination)
del android.zip
#endregion

#region list version of MDX
Write-host "$(Get-Date): List applications deployed on XMS"
$headers=@{"Content-Type" = "application/json"}
$Url = "https://${XMS}:4443/xenmobile/api/v1/authentication/login"
$json=Invoke-RestMethod -Uri $url -Body $cred -ContentType application/json -Headers $headers -Method POST -Verbose:$false
$headers.add("auth_token",$json.auth_token)
$mdx=
'
{ 
  "start": 0,
  "limit": 1000
}
'
$curmdx=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/application/filter" -Body $mdx -Headers $headers -Method Post
$count = $curmdx.applicationlistdata.applist.length
Write-host "$(Get-Date): Extract MDX applications from the list"
For($v=0;$v -lt $count;$v++)
{
if($curmdx.applicationlistdata.appList[$v].apptype -eq "MDX")
{
$curmdx.applicationlistdata.appList[$v].id >> curmdx.txt
}
}

#region Upgrade MDX
Write-host "$(Get-Date): Retrieve applications names and versions"
$test = get-content curmdx.txt
foreach($line in $test)
{
$headers=@{"Content-Type" = "application/json"}
$Url = "https://${XMS}:4443/xenmobile/api/v1/authentication/login"
$json=Invoke-RestMethod -Uri $url -Body $cred -ContentType application/json -Headers $headers -Method POST -Verbose:$false
$headers.add("auth_token",$json.auth_token)
$vermdx=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/application/mobile/$line" -Headers $headers -Method Get


#region Secure Mail
if(($vermdx.container.name -eq "Secure Mail") -and ($vermdx.container.ios.appversion))
{
$filename = "C:\Temp\Secure Apps\iOS\SecureMail-AppStore-10.7.10.22.mdx"
$platform = "ios"
$id = $vermdx.container.id
$description = $vermdx.container.description
$deliverygroups =  $vermdx.container.roles
$categories = $vermdx.container.categories
$appName = $vermdx.container.name
write-host "Update $appname for $platform"
Update-MDX
}
if(($vermdx.container.name -eq "Secure Mail") -and ($vermdx.container.android.appversion))
{
$filename = "C:\Temp\Secure Apps\Android\SecureMail-PlayStore-10.7.10.13.mdx"
$platform = "android"
$id = $vermdx.container.id
$description = $vermdx.container.description
$deliverygroups =  $vermdx.container.roles
$categories = $vermdx.container.categories
$appName = $vermdx.container.name
write-host "Update $appname for $platform"
Update-MDX
}
#endregion

#region Secure Web
if(($vermdx.container.name -eq "Secure Web") -and ($vermdx.container.ios.appversion))
{
$filename = "C:\Temp\Secure Apps\iOS\SecureWeb-AppStore-10.7.10.25.mdx"
$platform = "ios"
$id = $vermdx.container.id
$description = $vermdx.container.description
$deliverygroups =  $vermdx.container.roles
$categories = $vermdx.container.categories
$appName = $vermdx.container.name
write-host "Update $appname for $platform"
Update-MDX
}
if(($vermdx.container.name -eq "Secure Web") -and ($vermdx.container.android.appversion))
{
$filename = "C:\Temp\Secure Apps\Android\SecureWeb-Playstore-10.7.10.5.mdx"
$platform = "android"
$id = $vermdx.container.id
$description = $vermdx.container.description
$deliverygroups =  $vermdx.container.roles
$categories = $vermdx.container.categories
$appName = $vermdx.container.name
write-host "Update $appname for $platform"
Update-MDX
}
#endregion

#region Secure Notes
if(($vermdx.container.name -eq "Secure Notes") -and ($vermdx.container.ios.appversion))
{
$filename = "C:\Temp\Secure Apps\iOS\SecureNotes-AppStore-10.7.0.225.mdx"
$platform = "ios"
$id = $vermdx.container.id
$description = $vermdx.container.description
$deliverygroups =  $vermdx.container.roles
$categories = $vermdx.container.categories
$appName = $vermdx.container.name
write-host "Update $appname for $platform"
Update-MDX
}
if(($vermdx.container.name -eq "Secure Notes") -and ($vermdx.container.android.appversion))
{
$filename = "C:\Temp\Secure Apps\Android\SecureNotes-PlayStore-10.6.20.3.mdx"
$platform = "android"
$id = $vermdx.container.id
$description = $vermdx.container.description
$deliverygroups =  $vermdx.container.roles
$categories = $vermdx.container.categories
$appName = $vermdx.container.name
write-host "Update $appname for $platform"
Update-MDX
}
#endregion

#region Secure Tasks
if(($vermdx.container.name -eq "Secure Tasks") -and ($vermdx.container.ios.appversion))
{
$filename = "C:\Temp\Secure Apps\iOS\SecureTasks-AppStore-10.7.0.67.mdx"
$platform = "ios"
$id = $vermdx.container.id
$description = $vermdx.container.description
$deliverygroups =  $vermdx.container.roles
$categories = $vermdx.container.categories
$appName = $vermdx.container.name
write-host "Update $appname for $platform"
Update-MDX
}
if(($vermdx.container.name -eq "Secure Tasks") -and ($vermdx.container.android.appversion))
{
$filename = "C:\Temp\Secure Apps\Android\SecureTasks-Playstore-10.6.20.3.mdx"
$platform = "android"
$id = $vermdx.container.id
$description = $vermdx.container.description
$deliverygroups =  $vermdx.container.roles
$categories = $vermdx.container.categories
$appName = $vermdx.container.name
write-host "Update $appname for $platform"
Update-MDX
}
#endregion

#region QuickEdit
if(($vermdx.container.name -eq "QuickEdit") -and ($vermdx.container.ios.appversion))
{
$filename = "C:\Temp\Secure Apps\iOS\QuickEdit_AppStore_6.15_2.mdx"
$platform = "ios"
$id = $vermdx.container.id
$description = $vermdx.container.description
$deliverygroups =  $vermdx.container.roles
$categories = $vermdx.container.categories
$appName = $vermdx.container.name
write-host "Update $appname for $platform"
Update-MDX
}
if(($vermdx.container.name -eq "QuickEdit") -and ($vermdx.container.android.appversion))
{
$filename = "C:\Temp\Secure Apps\Android\QuickEdit_PlayStore_6.13.mdx"
$platform = "android"
$id = $vermdx.container.id
$description = $vermdx.container.description
$deliverygroups =  $vermdx.container.roles
$categories = $vermdx.container.categories
$appName = $vermdx.container.name
write-host "Update $appname for $platform"
Update-MDX
}
#endregion

#region ShareFile
if(($vermdx.container.name -match "Share") -and ($vermdx.container.ios.appversion))
{
$filename = "C:\Temp\Secure Apps\iOS\ShareFile_AppStore_6.3.mdx"
$platform = "ios"
$id = $vermdx.container.id
$description = $vermdx.container.description
$deliverygroups =  $vermdx.container.roles
$categories = $vermdx.container.categories
$appName = $vermdx.container.name
write-host "Update $appname for $platform"
Update-MDX
}
if(($vermdx.container.name -match "Share") -and ($vermdx.container.android.appversion))
{
$filename = "C:\Temp\Secure Apps\Android\ShareFile_PlayStore_6.0.mdx"
$platform = "android"
$id = $vermdx.container.id
$description = $vermdx.container.description
$deliverygroups =  $vermdx.container.roles
$categories = $vermdx.container.categories
$appName = $vermdx.container.name
write-host "Update $appname for $platform"
Update-MDX
}
#endregion
}
#endregion

#region end of script
CD \
$host.ui.RawUI.ForegroundColor = "Yellow"
write-host "Remove Temporary Folder and files"
$host.ui.RawUI.ForegroundColor = "White"
Remove-Item -path "C:\Temp" -recurse
#endregion

