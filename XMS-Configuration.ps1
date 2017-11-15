#Bypass certificate verification to enable access with XMS IP Address 
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

#Connect to XMS server 
$host.ui.RawUI.ForegroundColor = "White"
$Global:XMS = Read-Host -Prompt 'Please provide url of the XMS Server'

#Get Login credentials
write-host "Please provide username and password"
$Credential = get-credential $null

#Check Credentials before continue
$log = '{{"login":"{0}","password":"{1}"}}'
$Global:cred = ($log -f $Credential.UserName, $Credential.GetNetworkCredential().Password)

$headers=@{"Content-Type" = "application/json"}
$json=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/authentication/login -Body $cred -ContentType application/json -Headers $headers -Method POST
$headers.add("auth_token",$json.auth_token)

#region Licenses
$host.ui.RawUI.ForegroundColor = "yellow"
write-host "Configure Licenses"

$lic = 
'
{ 
  "serverAddress": "192.168.0.108", 
  "localPort": 0,
  "remotePort": 27000, 
  "serverPort": 8083, 
  "serverType": "remote", 
  "licenseType": "none", 
  "isServerConfigured": true, 
  "gracePeriodLeft": 0, 
  "isRestartLpeNeeded": true, 
  "isScheduleNotificationNeeded": true, 
  "licenseNotification": { 
    "id": 1, 
    "notificationEnabled": true, 
    "notifyFrequency": 2, 
    "notifyNumberDaysBeforeExpire": 10, 
    "recepientList": "arnaud.pain@arnaud.biz", 
    "emailContent": "XMS Licenses will expire soon" 
  }
}
'

$license = Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/licenses -Body $lic -ContentType application/json -Headers $headers -Method Post

$host.ui.RawUI.ForegroundColor = "white"
write-host "	Licenses Configured"
#endregion

#region Certificates

#endregion

#region Ldap
$host.ui.RawUI.ForegroundColor = "yellow"
write-host "Configure LDAP"

$payload = @{
	primaryHost='192.168.0.100'
	secondaryHost='192.168.0.100'
	port='389'
	username='xms-svc@arnaud.lab'
	password='Annec@r0le'
	userBaseDN='dc=arnaud,dc=lab'
	groupBaseDN='dc=arnaud,dc=lab'
	lockoutLimit='0'
	lockoutTime='1'
	useSecure='false'
	userSearchBy='samaccount'
	domain='arnaud.lab'
	domainAlias='arnaud'
	globalCatalogPort='3268'
	gcRootContext=''
}
$ad = $payload | ConvertTo-Json

$ldap = Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/ldap/msactivedirectory -Body $ad -ContentType application/json -Headers $headers -Method Post

$host.ui.RawUI.ForegroundColor = "white"
write-host "	LDAP Configured"
#endregion

#region NetScaler
$host.ui.RawUI.ForegroundColor = "yellow"
write-host "Configure NetScaler"

$payload = @{
 name='Display Name'
 alias='Alias'
 url='https://externalURL.com'
 passwordRequired='false'
 logonType='Domain'
 id='2'
 default='false'
}

$ns = $payload | ConvertTo-Json

$netscaler = Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/netscaler -Body $ns -ContentType application/json -Headers $headers -Method Post

$host.ui.RawUI.ForegroundColor = "white"
write-host "	NetScaler Configured"
#endregion

#region SMS Notification Server  
$host.ui.RawUI.ForegroundColor = "yellow"
write-host "Configure SMS Notification Server"
$host.ui.RawUI.ForegroundColor = "White"

$payload = @{
 name='displayName'
 description='Description'
 key='123456'
 secret='secreKey'
 virtualPhoneNumber='4086792222'
 https='false'
 country='+93'
 carrierGateway='true'
}

$notif = $payload | ConvertTo-Json

$smssrv = Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/notificationserver/sms -Body $notif -ContentType application/json -Headers $headers -Method Post

$host.ui.RawUI.ForegroundColor = "white"
write-host "	SMS Notification Server Configured"
#endregion

#region SMTP Notification Server
$host.ui.RawUI.ForegroundColor = "yellow"
write-host "Configure SMTP Notification Server"

$payload = @{
 name='SMTP Server 2'
 server='smtp.gmail.com'
 serverType='SMTP'
 description='SMTP Server'
 secureChannelProtocol='TLS'
 port='587'
 authentication='true'
 username='test@gmail.com'
 password='123'
 msSecurePasswordAuth='false'
 fromName='Test XMS'
 fromEmail='test@gail.com'
 numOfRetries='5'
 timeout='30'
 maxRecipients='100'
}

$smtp = $payload | ConvertTo-Json 

$smtpsrv = Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/notificationserver/smtp -Body $smtp -ContentType application/json -Headers $headers -Method Post

$host.ui.RawUI.ForegroundColor = "white"
write-host "	SMTP Notification Server Configured"
#endregion

#region New Client Properties
$host.ui.RawUI.ForegroundColor = "yellow"
write-host "Configure new client Properties"

$payload = @{
 	displayName='MyProperty'
	description='MyProperty Description'
	key='MyKey'
	value='15'
			}

$prop = $payload | ConvertTo-Json

$clientprop = Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/clientproperties -Body $prop -ContentType application/json -Headers $headers -Method Post

$host.ui.RawUI.ForegroundColor = "white"
write-host "	New Client Properties Configured"
#endregion

#region Update Client Properties
$host.ui.RawUI.ForegroundColor = "yellow"
write-host "Modify client Properties"

$payload = @{
 	displayName='Enable Citrix PIN Authentication'
	description='Enable Citrix PIN Authentication Description'
	value='true'
			}

$prop = $payload | ConvertTo-Json

$clientprop = Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/clientproperties/ENABLE_PASSCODE_AUTH -Body $prop -ContentType application/json -Headers $headers -Method Put

$host.ui.RawUI.ForegroundColor = "white"
write-host "	Client Properties Modified"
#endregion

#region Policy

#endregion

#region MDX Applications
$host.ui.RawUI.ForegroundColor = "yellow"
write-host "Add MDX Applications"

$xmhost=$XMS
$user=$Credential.UserName
$password=$Credential.GetNetworkCredential().Password
$filename="C:\E2EVC Barcelona\SecureMail_10.6.5.12.mdx"
$description="Uploaded by REST Services"
$platform="ios"
$deliverygroups="AllUsers"
$categories="Default"
$appname="Secure Mail"

#---------------------------------------------------------------------------------------------------------
#---	Functions
#---------------------------------------------------------------------------------------------------------

Function ConvertTo-PlainText( [security.securestring]$secure ) {
 $marshal = [Runtime.InteropServices.Marshal]
 $marshal::PtrToStringAuto( $marshal::SecureStringToBSTR($secure) )
}

#  Finding the correct URI to use...	
$BASEURI="https://" + $XMHOST + ":4443"

# If password has not been input through command line, let's ask for it in a secure way!
if ($password.Length -eq 0)
{
	$SecPass = Read-Host -assecurestring "Please enter your password"
	$PlainPass=ConvertTo-Plaintext($SecPass)
}
	else
{
	$PlainPass=$password
}

$count=0

$headers = @{
"Content-Type" = "application/json"
}

$json = "{
		""login"":""" + $User + """,
		""password"":""" + $PlainPass + """
		}"

#Get credential token
$URI=$BASEURI + "/xenmobile/api/v1/authentication/login"

$authtoken=try{
	Invoke-RestMethod -Headers $headers -Method Post -Uri $URI -Body $json
	} catch { $_.Exception.Response }

	if ($authtoken.StatusCode -ne $null)
	{
	if ($authtoken.StatusCode.Value__.Equals(401) -or $authtoken.StatusCode.Value__.Equals(400))
	{
		Write-Output "Bad username or password, unauthorized!"
		Exit
	}
}

$token=$authtoken.auth_token

$headers=@{"Auth_token"=$token
			"Content-type"="multipart/form-data; boundary=----XenmobileScripting----"
			"Accept"= "*/*"
			"Accept-Encoding"="gzip, deflate"
			"Accept-Language"="en-US,en;q=0.8\r\n"
}

$URI=$BASEURI + "/xenmobile/api/v1/application/mobile/mdx/" + $platform.ToLower()

$fileBin=[IO.File]::ReadAllBytes($filename)
$enc=[System.Text.Encoding]::GetEncoding("ISO-8859-1")
$fileenc=$enc.GetString($fileBin)

#Create a correctly formed DeliveryGroups object...
$deliverygroup=""
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
#$body
$body=$body + '------XenmobileScripting----
Content-Disposition: form-data;name="uploadFile";filename="' + (Get-Item $filename).basename + (Get-Item $filename).Extension + '"
Content-Type: application/octet-stream

' + $fileenc + '
------XenmobileScripting------
'
$host.ui.RawUI.ForegroundColor = "white"

$return=try{
	Invoke-RestMethod -Headers $headers -Method Post -Uri $URI -Body $body
	} catch { $_.Exception.Response }
	
	if($return.StatusCode -ne $null)
	{
		Write-Output "Error! Errormessage:" $return.Message "`r`n" $return.InnerException
	}
	else
	{
		Write-Output "	MDX Apps added"
	}
$host.ui.RawUI.ForegroundColor = "White"
#endregion

#region Action

#endregion

#region Delivery Groups
$host.ui.RawUI.ForegroundColor = "yellow"
write-host "Add Delivery Groups"

$dgroups = 
'
{ 
  "name": "DG1", 
  "description": "Delivery Group 1 Description", 
  "applications": [ 
    { 
      "name": "Secure Mail", 
      "priority": -1, 
      "required": false
    }
  ],
  "groups": [ 
    { 
    "name": "CN=XenMobile,CN=Users,DC=mtrc,DC=lab",
	"uniqueId": "XenMobile",
	"domainName": "mtrc.lab"
	}
  ]
}
}
'

$license = Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/deliverygroups -Body $dgroups -ContentType application/json -Headers $headers -Method Post

$host.ui.RawUI.ForegroundColor = "white"
write-host "	Delivery Groups added"
#endregion

#region RBAC
$host.ui.RawUI.ForegroundColor = "yellow"
write-host "Add RBAC"

$rbac = 
'
{
  "name": "ADMIN_11",
  "permissions": [
    { 
      "permission": "perm-feature-DEVICE-", 
      "granted": true 
    },
    { 
      "permission": "perm-feature-DEVICE_EDIT_PROPERTIES-", 
      "granted": true 
    }, 
    { 
      "permission": "perm-feature-DEVICE_EDIT-", 
      "granted": true
    }, 
    { 
      "permission": "perm-feature-SETTING-", 
      "granted": true
    }
  ], 
  "adGroups": [ 
    { 
      "primaryGroupToken": 545, 
      "uniqueName": "Users", 
      "domainName": "mtrc.lab"
    }
  ]
}
'

$rbacconf = Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/rbac/role -Body $rbac -ContentType application/json -Headers $headers -Method Post

$host.ui.RawUI.ForegroundColor = "white"
write-host "	RBAC added"
#endregion
