#Bypass certificate verification to enable access with XMS IP Address 
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

#Connect to XMS server 
$host.ui.RawUI.ForegroundColor = "White"
$Global:XMSource = Read-Host -Prompt 'Please provide url of the source XMS Server'

#Get Login credentials
write-host "Please provide username and password for Source XMS"
$Global:SourceCred = get-credential $null

$Global:XMSDest = Read-Host -Prompt 'Please provide url of the destination XMS Server'
write-host "Please provide username and password for Destination XMS"
$Global:DestCred = get-credential $null

#Check Credentials before continue
$logs = '{{"login":"{0}","password":"{1}"}}'
$Global:Scred = ($logs -f $SourceCred.UserName, $SourceCred.GetNetworkCredential().Password)
$logd = '{{"login":"{0}","password":"{1}"}}'
$Global:Dcred = ($logd -f $DestCred.UserName, $DestCred.GetNetworkCredential().Password)

#Check Credentials on Source XMS before continue
$headers=@{"Content-Type" = "application/json"}
$json=Invoke-RestMethod -Uri https://${XMSource}:4443/xenmobile/api/v1/authentication/login -Body $Scred -Headers $headers -Method POST
$headers.add("auth_token",$json.auth_token)

#Check Credentials on Destination XMS before continue
$headers=@{"Content-Type" = "application/json"}
$json=Invoke-RestMethod -Uri https://${XMSDest}:4443/xenmobile/api/v1/authentication/login -Body $Dcred -Headers $headers -Method POST
$headers.add("auth_token",$json.auth_token)

#region Licenses

# Export Licenses Information for source host
$headers=@{"Content-Type" = "application/json"}
$json=Invoke-RestMethod -Uri https://${XMSource}:4443/xenmobile/api/v1/authentication/login -Body $Scred  -Headers $headers -Method POST
$headers.add("auth_token",$json.auth_token)

$license=Invoke-RestMethod -Uri "https://${XMSource}:4443/xenmobile/api/v1/licenses" -Headers $headers -Method Get

$host.ui.RawUI.ForegroundColor = "Yellow"
Write-host "$(Get-Date): Export Licenses configuration"
$host.ui.RawUI.ForegroundColor = "white"

# Declare an array to collect our result objects
$resultsarray =@()

# For every $contact held in the $contacts, do this loop
foreach($cpLicenseServer in $license){

# Create a new custom object to hold our result.
$contactObject = new-object PSObject

# Add our data to $contactObject as attributes using the add-member commandlet
$contactObject | add-member -membertype NoteProperty -name serveraddress -Value $cpLicenseServer.cpLicenseServer.serveraddress
$contactObject | add-member -membertype NoteProperty -name localport -Value $cpLicenseServer.cpLicenseServer.localport
$contactObject | add-member -membertype NoteProperty -name remoteport -Value $cpLicenseServer.cpLicenseServer.remoteport
$contactObject | add-member -membertype NoteProperty -name serverport -Value $cpLicenseServer.cpLicenseServer.serverport
$contactObject | add-member -membertype NoteProperty -name servertype -Value $cpLicenseServer.cpLicenseServer.servertype
$contactObject | add-member -membertype NoteProperty -name licenseType -Value $cpLicenseServer.cpLicenseServer.licenseType
$contactObject | add-member -membertype NoteProperty -name isserverconfigured -Value $cpLicenseServer.cpLicenseServer.isserverconfigured
$contactObject | add-member -membertype NoteProperty -name graceperiodleft -Value $cpLicenseServer.cpLicenseServer.graceperiodleft
$contactObject | add-member -membertype NoteProperty -name isrestartedlpeneeded -Value $cpLicenseServer.cpLicenseServer.isrestartedlpeneeded
$contactObject | add-member -membertype NoteProperty -name isshcedulednotificationneeded -Value $cpLicenseServer.cpLicenseServer.isshcedulednotificationneeded
$contactObject | add-member -membertype NoteProperty -name id -Value $cpLicenseServer.cpLicenseServer.licensenotification.id
$contactObject | add-member -membertype NoteProperty -name notificationenabled -Value $cpLicenseServer.cpLicenseServer.licensenotification.notificationenabled
$contactObject | add-member -membertype NoteProperty -name notifyfrequency -Value $cpLicenseServer.cpLicenseServer.licensenotification.notifyfrequency
$contactObject | add-member -membertype NoteProperty -name notifynumberdaysbeforeexpire -Value $cpLicenseServer.cpLicenseServer.licensenotification.notifynumberdaysbeforeexpire
$contactObject | add-member -membertype NoteProperty -name recepientlist -Value $cpLicenseServer.cpLicenseServer.licensenotification.recepientlist
$contactObject | add-member -membertype NoteProperty -name emailcontent -Value $cpLicenseServer.cpLicenseServer.licensenotification.emailcontent

# Save the current $contactObject by appending it to $resultsArray ( += means append a new element to ‘me’)
$resultsarray += $contactObject
}

$resultsarray| Export-csv "Licenses.csv" -delimiter ";"

$host.ui.RawUI.ForegroundColor = "yellow"
write-host "$(Get-Date):  Import Licenses configuration"

$data = (get-content "licenses.csv")[2].split(";") |ForEach-Object {$_ -replace '"', ''}
foreach($value in $data)
{
$lic = 
'
{ 
  "serverAddress": "' + $data[0] +'", 
  "localPort": "' + $data[1] +'",
  "remotePort": "' + $data[2] +'", 
  "serverPort": "' + $data[3] +'", 
  "serverType": "' + $data[4] +'", 
  "licenseType": "' + $data[5] +'", 
  "isServerConfigured": "' + $data[6] +'", 
  "gracePeriodLeft": "' + $data[7] +'", 
  "isRestartLpeNeeded": "true", 
  "isScheduleNotificationNeeded": true, 
  "licenseNotification": { 
    "id": "' + $data[10] +'", 
    "notificationEnabled": "' + $data[11] +'", 
    "notifyFrequency": "' + $data[12] +'", 
    "notifyNumberDaysBeforeExpire": "' + $data[13] +'", 
    "recepientList": "' + $data[14] +'", 
    "emailContent": "' + $data[15] +'"
  }
  }
}
'
}
$headers=@{"Content-Type" = "application/json"}
$json=Invoke-RestMethod -Uri https://${XMSDest}:4443/xenmobile/api/v1/authentication/login -Body $Dcred  -Headers $headers -Method POST
$headers.add("auth_token",$json.auth_token)
$license = Invoke-RestMethod -Uri https://${XMSDest}:4443/xenmobile/api/v1/licenses -Body $lic -Headers $headers -Method Post

$host.ui.RawUI.ForegroundColor = "Green"
write-host "$(Get-Date):	Licenses Configured"
#endregion

#region Ldap
# Export LDAP Information for source host
$headers=@{"Content-Type" = "application/json"}
$json=Invoke-RestMethod -Uri https://${XMSource}:4443/xenmobile/api/v1/authentication/login -Body $Scred -Headers $headers -Method POST
$headers.add("auth_token",$json.auth_token)

$ldap=Invoke-RestMethod -Uri "https://${XMSource}:4443/xenmobile/api/v1/ldap" -Headers $headers -Method Get
$host.ui.RawUI.ForegroundColor = "Yellow"
Write-host "$(Get-Date): Extract LDAP configuration"
$host.ui.RawUI.ForegroundColor = "white"

# Declare an array to collect our result objects
$resultsarray =@()

# For every $contact held in the $contacts, do this loop
foreach($adlist in $ldap)
{

# Create a new custom object to hold our result.
$contactObject = new-object PSObject

# Add our data to $contactObject as attributes using the add-member commandlet
$contactObject | add-member -membertype NoteProperty -name "Server Address" -Value $adlist.adlist.domain
$contactObject | add-member -membertype NoteProperty -name "User Base DN" -Value $adlist.adlist.userbasedn
$contactObject | add-member -membertype NoteProperty -name "Group Base DN" -Value $adlist.adlist.groupbasedn
$contactObject | add-member -membertype NoteProperty -name "Port" -Value $adlist.adlist.port
$contactObject | add-member -membertype NoteProperty -name "Username" -Value $adlist.adlist.username
$contactObject | add-member -membertype NoteProperty -name "Primary Host" -Value $adlist.adlist.primaryhost
$contactObject | add-member -membertype NoteProperty -name "Secondary Host" -Value $adlist.adlist.secondaryhost
$contactObject | add-member -membertype NoteProperty -name "Use Secure" -Value $adlist.adlist.usesecure
$contactObject | add-member -membertype NoteProperty -name "Global Catalog Port" -Value $adlist.adlist.globalcatalogport
$contactObject | add-member -membertype NoteProperty -name "User Search By" -Value $adlist.adlist.usersearchby
$contactObject | add-member -membertype NoteProperty -name "Domain Alias" -Value $adlist.adlist.domainalias

# Save the current $contactObject by appending it to $resultsArray ( += means append a new element to ‘me’)
$resultsarray += $contactObject
}

$resultsarray| Export-csv "ldap.csv" -delimiter ";"

$host.ui.RawUI.ForegroundColor = "yellow"
write-host "$(Get-Date):  Import LDAP configuration"

$host.ui.RawUI.ForegroundColor = "Red"
$pwd = Read-Host -Prompt 'Please provide password for LDAP Account' 

$data = (get-content "ldap.csv")[2].split(";") |ForEach-Object {$_ -replace '"', ''}
foreach($value in $data)
{
$payload = @{
	primaryHost=$data[5]
	secondaryHost=$data[6]
	port=$data[3]
	username=$data[4]
	password=$pwd
	userBaseDN=$data[2]
	groupBaseDN=$data[2]
	lockoutLimit='0'
	lockoutTime='1'
	useSecure=$data[7]
	userSearchBy=$data[9]
	domain=$data[0]
	domainAlias=$data[10]
	globalCatalogPort=$data[8]
	gcRootContext=''
}
}
$ad = $payload | ConvertTo-Json

$headers=@{"Content-Type" = "application/json"}
$json=Invoke-RestMethod -Uri https://${XMSDest}:4443/xenmobile/api/v1/authentication/login -Body $Dcred -Headers $headers -Method POST
$headers.add("auth_token",$json.auth_token)
$ldap = Invoke-RestMethod -Uri https://${XMSDest}:4443/xenmobile/api/v1/ldap/msactivedirectory -Body $ad -Headers $headers -Method Post

$host.ui.RawUI.ForegroundColor = "Green"
write-host "$(Get-Date):	LDAP Configured"
#endregion

#region NetScaler
# Export NetScaler Information for source host
$headers=@{"Content-Type" = "application/json"}
$json=Invoke-RestMethod -Uri https://${XMSource}:4443/xenmobile/api/v1/authentication/login -Body $Scred -Headers $headers -Method POST
$headers.add("auth_token",$json.auth_token)
$netscaler=Invoke-RestMethod -Uri "https://${XMSource}:4443/xenmobile/api/v1/netscaler" -Headers $headers -Method Get

$host.ui.RawUI.ForegroundColor = "Yellow"
Write-host "$(Get-Date): Export NetScaler configuration"
$host.ui.RawUI.ForegroundColor = "white"

# Declare an array to collect our result objects
$resultsarray =@()

# For every $contact held in the $contacts, do this loop
foreach($aglist in $netscaler)
{

# Create a new custom object to hold our result.
$contactObject = new-object PSObject

# Add our data to $contactObject as attributes using the add-member commandlet
$contactObject | add-member -membertype NoteProperty -name "Name" -Value $aglist.aglist.name
$contactObject | add-member -membertype NoteProperty -name "Alias" -Value $aglist.aglist.alias
$contactObject | add-member -membertype NoteProperty -name "URL" -Value $aglist.aglist.url
$contactObject | add-member -membertype NoteProperty -name "Password Required" -Value $aglist.aglist.passwordrequired
$contactObject | add-member -membertype NoteProperty -name "Logon Type" -Value $aglist.aglist.logonType
$contactObject | add-member -membertype NoteProperty -name "Callback" -Value $aglist.aglist.callback
$contactObject | add-member -membertype NoteProperty -name "ID" -Value $aglist.aglist.id
$contactObject | add-member -membertype NoteProperty -name "Default" -Value $aglist.aglist.default

# Save the current $contactObject by appending it to $resultsArray ( += means append a new element to ‘me’)
$resultsarray += $contactObject
}
$resultsarray| Export-csv "netscaler.csv" -delimiter ";"

# Import NetScaler Information on destination host
$host.ui.RawUI.ForegroundColor = "yellow"
write-host "$(Get-Date):  Import NetScaler configuration"
$data = (get-content "netscaler.csv")[2].split(";") | ForEach-Object {$_ -replace '"', ''}
if($data[3] -eq "True"){$data[3]="false"}
if($data[3] -eq "False"){$data[3]="true"}
foreach($value in $data)
{
$ns = 
'
{ 
  "name": "' + $data[0] +'", 
  "alias": "' + $data[1] +'",
  "url": "' + $data[2] +'", 
  "passwordRequired": ' + $data[3] +', 
  "logonType": "' + $data[4] +'", 
  "default": ' + $data[6] +'
  }
'
}

$headers=@{"Content-Type" = "application/json"}
$json=Invoke-RestMethod -Uri https://${XMSDest}:4443/xenmobile/api/v1/authentication/login -Body $Dcred -Headers $headers -Method POST
$headers.add("auth_token",$json.auth_token)
$netscaler = Invoke-RestMethod -Uri https://${XMSDest}:4443/xenmobile/api/v1/netscaler -Body $ns -Headers $headers -Method Post

$host.ui.RawUI.ForegroundColor = "Green"
write-host "$(Get-Date):	NetScaler Configured"
#endregion

#region Notification Server  
$headers=@{"Content-Type" = "application/json"}
$json=Invoke-RestMethod -Uri https://${XMSource}:4443/xenmobile/api/v1/authentication/login -Body $Scred -Headers $headers -Method POST
$headers.add("auth_token",$json.auth_token)
$notification=Invoke-RestMethod -Uri "https://${XMSource}:4443/xenmobile/api/v1/notificationserver" -Headers $headers -Method Get

$host.ui.RawUI.ForegroundColor = "Yellow"
Write-host "$(Get-Date): Export Notification Server configuration"
$host.ui.RawUI.ForegroundColor = "white"

# Declare an array to collect our result objects
$resultsarray =@()

# For every $contact held in the $contacts, do this loop
$new=$number-$number
$result=$new
$count = $notification.list.length
if($count -gt 1)
{
for ($v=0;$v -lt $count; $v++)
{foreach ($list in $notification)
{
# Create a new custom object to hold our result.
$contactObject = new-object PSObject
# Add our data to $contactObject as attributes using the add-member commandlet
$contactObject | add-member -membertype NoteProperty -name "ServerType" -Value $list.list.serverType[$new] 
$contactObject | add-member -membertype NoteProperty -name "ID" -Value $list.list.id[$new]
$new++
# Save the current $contactObject by appending it to $resultsArray ( += means append a new element to ‘me’)
$resultsarray += $contactObject
}
}
$resultsarray| Export-csv "id.csv" -delimiter ","
}

#Retreive Type of server
$file = import-csv "id.csv"
foreach($line in $file)
{
 if($line.ServerType -eq "SMTP")
 {
 	$SMTP = import-csv "id.csv"
	foreach ($line in $SMTP)
		{
		if($line.ServerType -eq "SMTP"){$id = $line.id}
		}
	#export data
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMSource}:4443/xenmobile/api/v1/authentication/login -Body $Scred -Headers $headers -Method POST
	$headers.add("auth_token",$json.auth_token)	
	$ssmtp=Invoke-RestMethod -Uri "https://${XMSource}:4443/xenmobile/api/v1/notificationserver/$id" -Headers $headers -Method Get
	# Declare an array to collect our result objects
	$resultsarray =@()
	# Create a new custom object to hold our result.
	$contactObject = new-object PSObject

	# Add our data to $contactObject as attributes using the add-member commandlet
	$contactObject | add-member -membertype NoteProperty -name "id" -Value $ssmtp.details.id
	$contactObject | add-member -membertype NoteProperty -name "active" -Value $ssmtp.details.active
	$contactObject | add-member -membertype NoteProperty -name "name" -Value $ssmtp.details.name
	$contactObject | add-member -membertype NoteProperty -name "server" -Value $ssmtp.details.server
	$contactObject | add-member -membertype NoteProperty -name "servertype" -Value $ssmtp.details.servertype
	$contactObject | add-member -membertype NoteProperty -name "description" -Value $ssmtp.details.description
	$contactObject | add-member -membertype NoteProperty -name "secureChannelProtocol" -Value $ssmtp.details.secureChannelProtocol
	$contactObject | add-member -membertype NoteProperty -name "port" -Value $ssmtp.details.port
	$contactObject | add-member -membertype NoteProperty -name "authentication" -Value $ssmtp.details.authentication
	$contactObject | add-member -membertype NoteProperty -name "username" -Value $ssmtp.details.username
	$contactObject | add-member -membertype NoteProperty -name "password" -Value $ssmtp.details.password
	$contactObject | add-member -membertype NoteProperty -name "msSecurePasswordAuth" -Value $ssmtp.details.msSecurePasswordAuth
	$contactObject | add-member -membertype NoteProperty -name "fromName" -Value $ssmtp.details.fromName
	$contactObject | add-member -membertype NoteProperty -name "fromEmail" -Value $ssmtp.details.fromEmail
	$contactObject | add-member -membertype NoteProperty -name "numOfRetries" -Value $ssmtp.details.numOfRetries
	$contactObject | add-member -membertype NoteProperty -name "timeout" -Value $ssmtp.details.timeout
	$contactObject | add-member -membertype NoteProperty -name "maxRecipients" -Value $ssmtp.details.maxRecipients
	# Save the current $contactObject by appending it to $resultsArray ( += means append a new element to ‘me’)
	$resultsarray += $contactObject
	$resultsarray| Export-csv "$id.csv" -delimiter ";"

	$host.ui.RawUI.ForegroundColor = "yellow"
	write-host "$(Get-Date):  Import SMTP Notification Server configuration"
	
	#Import SMTP
	$Import = (get-content "$id.csv")[2].split(";") | ForEach-Object {$_ -replace '"', ''}
		if($Import[8] -eq "False"){$Import[8]="false"}
		if($Import[11] -eq "False"){$Import[11]="false"}
		if($Import -eq "SMTP")
			{
			$smtp = 
			'
			{ 
			"name": "' + $Import[2] +'",
			"server": "' + $Import[3] +'",
			"serverType": "' + $Import[4] +'",
			"description": "' + $Import[5] +'",
			"secureChannelProtocol": "' + $Import[6] +'",
			"port": ' + $Import[7] +',
			"authentication": ' + $Import[8] +',
			"username": "' + $Import[9] +'",
			"password": "' + $Import[10] +'", 
			"msSecurePasswordAuth": ' + $Import[11] +',
			"fromName": "' + $Import[12] +'",
			"fromEmail": "' + $Import[13] +'",
			"numOfRetries": ' + $Import[14] +',
			"timeout": ' + $Import[15] +',
			"maxRecipients": ' + $Import[16] +'
			}
			'
			}
	
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMSDest}:4443/xenmobile/api/v1/authentication/login -Body $Dcred -Headers $headers -Method POST
	$headers.add("auth_token",$json.auth_token)
	$smtpsrv = Invoke-RestMethod -Uri https://${XMSDest}:4443/xenmobile/api/v1/notificationserver/smtp -Body $smtp -Headers $headers -Method Post
	$host.ui.RawUI.ForegroundColor = "Green"
	write-host "$(Get-Date):	SMTP Notification Server Configured"
 }
	 
 elseif($line.ServerType -eq "SMS")
 {
	$host.ui.RawUI.ForegroundColor = "yellow"
	write-host "$(Get-Date):  Import SMS Notification Server configuration"
	$SMS = import-csv "id.csv"
	foreach ($line in $SMS)
		{
		if($line.ServerType -eq "SMS"){$id = $line.id}
		}
	#export data
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMSource}:4443/xenmobile/api/v1/authentication/login -Body $Scred -Headers $headers -Method POST
	$headers.add("auth_token",$json.auth_token)
	$ssms=Invoke-RestMethod -Uri "https://${XMSource}:4443/xenmobile/api/v1/notificationserver/$id" -Headers $headers -Method Get
	# Declare an array to collect our result objects
	$resultsarray =@()
	# Create a new custom object to hold our result.
	$contactObject = new-object PSObject

	# Add our data to $contactObject as attributes using the add-member commandlet
	$contactObject | add-member -membertype NoteProperty -name "id" -Value $ssms.details.id
	$contactObject | add-member -membertype NoteProperty -name "active" -Value $ssms.details.active
	$contactObject | add-member -membertype NoteProperty -name "name" -Value $ssms.details.name
	$contactObject | add-member -membertype NoteProperty -name "server" -Value $ssms.details.server
	$contactObject | add-member -membertype NoteProperty -name "servertype" -Value $ssms.details.servertype
	$contactObject | add-member -membertype NoteProperty -name "description" -Value $ssms.details.description
	$contactObject | add-member -membertype NoteProperty -name "key" -Value $ssms.details.key
	$contactObject | add-member -membertype NoteProperty -name "secret" -Value $ssms.details.secret
	$contactObject | add-member -membertype NoteProperty -name "virtualPhoneNumber" -Value $ssms.details.virtualPhoneNumber
	$contactObject | add-member -membertype NoteProperty -name "https" -Value $ssms.details.https
	$contactObject | add-member -membertype NoteProperty -name "country" -Value $ssms.details.country
	$contactObject | add-member -membertype NoteProperty -name "carrierGateway" -Value $ssms.details.carrierGateway
	# Save the current $contactObject by appending it to $resultsArray ( += means append a new element to ‘me’)
	$resultsarray += $contactObject
	$resultsarray| Export-csv "$id.csv" -delimiter ";"
	
	#Import SMS
	$Import = (get-content "$id.csv")[2].split(";") | ForEach-Object {$_ -replace '"', ''}
		if($Import -eq "SMS")
			{
			$sms = 
			'
			{ 
			"name": "' + $Import[2] +'",
			"description": "' + $Import[5] +'",
			"key": "' + $Import[6] +'",
			"secret": "' + $Import[7] +'",
			"virtualPhoneNumber": "' + $Import[8] +'",
			"https": "' + $Import[9] +'",
			"country": "' + $Import[10] +'",
			"carrierGateway": "' + $Import[11] +'"
			}
			'
			}
			
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMSDest}:4443/xenmobile/api/v1/authentication/login -Body $Dcred -Headers $headers -Method POST
	$headers.add("auth_token",$json.auth_token)
	$smsrv = Invoke-RestMethod -Uri https://${XMSDest}:4443/xenmobile/api/v1/notificationserver/sms -Body $sms -Headers $headers -Method Post
	$host.ui.RawUI.ForegroundColor = "Green"
	write-host "$(Get-Date):	SMS Notification Server Configured"
 }

$host.ui.RawUI.ForegroundColor = "White"
}
#endregion

#region Client Properties
#region Export
$headers=@{"Content-Type" = "application/json"}
$json=Invoke-RestMethod -Uri https://${XMSource}:4443/xenmobile/api/v1/authentication/login -Body $Scred -Headers $headers -Method POST
$headers.add("auth_token",$json.auth_token)
$clientprop=Invoke-RestMethod -Uri "https://${XMSource}:4443/xenmobile/api/v1/clientproperties" -Headers $headers -Method Get

$host.ui.RawUI.ForegroundColor = "Yellow"
Write-host "$(Get-Date): Export Client Properties"
$host.ui.RawUI.ForegroundColor = "white"

# Declare an array to collect our result objects
$resultsarray =@()

# For every $count do this loop
$new=0
$count = $clientprop.allclientproperties.length
for ($v=0; $v -lt $count; $v++)
{
foreach($allclientproperties in $clientprop)
{
# Create a new custom object to hold our result.
$contactObject = new-object PSObject
# Add our data to $contactObject as attributes using the add-member commandlet
$contactObject | add-member -membertype NoteProperty -name "Display name" -Value $allclientproperties.allclientproperties[$new].displayname
$contactObject | add-member -membertype NoteProperty -name "Description" -Value $allclientproperties.allclientproperties[$new].description
$contactObject | add-member -membertype NoteProperty -name "Key" -Value $allclientproperties.allclientproperties[$new].key
$contactObject | add-member -membertype NoteProperty -name "Value" -Value $allclientproperties.allclientproperties[$new].value
}
$new++
# Save the current $contactObject by appending it to $resultsArray ( += means append a new element to ‘me’)
$resultsarray += $contactObject
}
$resultsarray| Export-csv "clientprop.csv" -delimiter ";"
#endregion Export

$host.ui.RawUI.ForegroundColor = "yellow"
write-host "$(Get-Date):  Update Client Properties configuration"

#region Update Client Properties 
for ($y=0; $y -lt $count; $y++)
{
$import = (get-content "clientprop.csv")[$y].split(";") |ForEach-Object {$_ -replace '"', ''}
	$clientprop = 
		'
		{ 
		"displayName": "' + $Import[0] +'",
		"description": "' + $Import[1] +'",
		"value": "' + $Import[3] +'"
		}
		'
		$key=$Import[2]
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMSDest}:4443/xenmobile/api/v1/authentication/login -Body $Dcred -Headers $headers -Method POST
	$headers.add("auth_token",$json.auth_token)
	$clientp = Invoke-RestMethod -Uri https://${XMSDest}:4443/xenmobile/api/v1/clientproperties/$key -Body $clientprop -Headers $headers -Method Put
}
#endregion Update

#region Add New Client Properties 
$host.ui.RawUI.ForegroundColor = "yellow"
write-host "$(Get-Date):  Configure New Client Properties configuration"
$headers=@{"Content-Type" = "application/json"}
$json=Invoke-RestMethod -Uri https://${XMSource}:4443/xenmobile/api/v1/authentication/login -Body $Scred -Headers $headers -Method POST
$headers.add("auth_token",$json.auth_token)
$clientprop=Invoke-RestMethod -Uri "https://${XMSource}:4443/xenmobile/api/v1/clientproperties" -Headers $headers -Method Get
$count = $clientprop.allclientproperties.length
for ($y=21; $y -lt $count+2; $y++)
{
$import = (get-content "clientprop.csv")[$y].split(";") |ForEach-Object {$_ -replace '"', ''}
	$newcp = 
		'
		{ 
		"displayName": "' + $Import[0] +'",
		"description": "' + $Import[1] +'",
		"key": "' + $Import[2] +'",
		"value": "' + $Import[3] +'"
		}
		'
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMSDest}:4443/xenmobile/api/v1/authentication/login -Body $Dcred -Headers $headers -Method POST
	$headers.add("auth_token",$json.auth_token)
	$clientp = Invoke-RestMethod -Uri https://${XMSDest}:4443/xenmobile/api/v1/clientproperties -Body $newcp -Headers $headers -Method Post
}
#endregion

$host.ui.RawUI.ForegroundColor = "Green"
write-host "$(Get-Date):	Client Properties Updated"
$host.ui.RawUI.ForegroundColor = "White"
#endregion Client Properties


$host.ui.RawUI.ForegroundColor = "White"

