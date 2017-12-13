#Requires -Version 3.0
#This File is in Unicode format.  Do not edit in an ASCII editor.

#region help text

<#
.SYNOPSIS
	Creates a XenMobile Configuration report.
.DESCRIPTION
	Creates a XenMobile Configuration report using Microsoft Word, PDF, formatted text or HTML and PowerShell.
	This script is based on Carl Webster Template Script and user XenMobile REST API.
	Creates a document named XMS-Report.docx (or .PDF or .TXT or .HTML).
	Word and PDF documents include a Cover Page, Table of Contents and Footer.
	Includes support for the following language versions of Microsoft Word:
		Catalan
		Chinese
		Danish
		Dutch
		English
		Finnish
		French
		German
		Norwegian
		Portuguese
		Spanish
		Swedish

.PARAMETER CompanyName
	Company Name to use for the Cover Page.  
	Default value is contained in HKCU:\Software\Microsoft\Office\Common\UserInfo\CompanyName or
	HKCU:\Software\Microsoft\Office\Common\UserInfo\Company, whichever is populated on the 
	computer running the script.
	This parameter has an alias of CN.
	If either registry key does not exist and this parameter is not specified, the report will
	not contain a Company Name on the cover page.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER CompanyAddress
	Company Address to use for the Cover Page, if the Cover Page has the Address field.  
		The following Cover Pages have an Address field:
			Sideline
			Contrast
			Exposure
			Filigree
			Ion (Dark)
			Retrospect
			Semaphore
			Tiles
			ViewMaster
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CA.
.PARAMETER CompanyEmail
	Company Email to use for the Cover Page, if the Cover Page has the Email field.  
		The following Cover Pages have an Email field:
			Facet
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CE.
.PARAMETER CompanyFax
	Company Fax to use for the Cover Page, if the Cover Page has the Fax field.  
		The following Cover Pages have a Fax field:
			Contrast
			Exposure
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CF.
.PARAMETER CompanyPhone
	Company Phone to use for the Cover Page, if the Cover Page has the Phone field.  
		The following Cover Pages have a Phone field:
			Contrast
			Exposure
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CPh.
.PARAMETER CoverPage
	What Microsoft Word Cover Page to use.
	Only Word 2010, 2013 and 2016 are supported.
	(default cover pages in Word en-US)
	
	Valid input is:
		Alphabet (Word 2010. Works)
		Annual (Word 2010. Doesn't work well for this report)
		Austere (Word 2010. Works)
		Austin (Word 2010/2013/2016. Doesn't work in 2013 or 2016, mostly works in 2010 but 
						Subtitle/Subject & Author fields need to be moved 
						after title box is moved up)
		Banded (Word 2013/2016. Works)
		Conservative (Word 2010. Works)
		Contrast (Word 2010. Works)
		Cubicles (Word 2010. Works)
		Exposure (Word 2010. Works if you like looking sideways)
		Facet (Word 2013/2016. Works)
		Filigree (Word 2013/2016. Works)
		Grid (Word 2010/2013/2016. Works in 2010)
		Integral (Word 2013/2016. Works)
		Ion (Dark) (Word 2013/2016. Top date doesn't fit; box needs to be manually resized or font 
						changed to 8 point)
		Ion (Light) (Word 2013/2016. Top date doesn't fit; box needs to be manually resized or font 
						changed to 8 point)
		Mod (Word 2010. Works)
		Motion (Word 2010/2013/2016. Works if top date is manually changed to 36 point)
		Newsprint (Word 2010. Works but date is not populated)
		Perspective (Word 2010. Works)
		Pinstripes (Word 2010. Works)
		Puzzle (Word 2010. Top date doesn't fit; box needs to be manually resized or font 
					changed to 14 point)
		Retrospect (Word 2013/2016. Works)
		Semaphore (Word 2013/2016. Works)
		Sideline (Word 2010/2013/2016. Doesn't work in 2013 or 2016, works in 2010)
		Slice (Dark) (Word 2013/2016. Doesn't work)
		Slice (Light) (Word 2013/2016. Doesn't work)
		Stacks (Word 2010. Works)
		Tiles (Word 2010. Date doesn't fit unless changed to 26 point)
		Transcend (Word 2010. Works)
		ViewMaster (Word 2013/2016. Works)
		Whisp (Word 2013/2016. Works)
		
	Default value is Sideline.
	This parameter has an alias of CP.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER UserName
	User name to use for the Cover Page and Footer.
	Default value is contained in $env:username
	This parameter has an alias of UN.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER PDF
	SaveAs PDF file instead of DOCX file.
	This parameter is disabled by default.
	The PDF file is roughly 5X to 10X larger than the DOCX file.
	This parameter requires Microsoft Word to be installed.
	This parameter uses the Word SaveAs PDF capability.
.PARAMETER Text
	Creates a formatted text file with a .txt extension.
	This parameter is disabled by default.
.PARAMETER MSWord
	SaveAs DOCX file
	This parameter is set True if no other output format is selected.
.PARAMETER HTML
	Creates an HTML file with an .html extension.
	This parameter is disabled by default.
	This parameter is reserved for a future update and no output is created at this time.
.PARAMETER AddDateTime
	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2017 at 6PM is 2017-06-01_1800.
	Output filename will be ReportName_2017-06-01_1800.docx (or .pdf).
	This parameter is disabled by default.
.PARAMETER Folder
	Specifies the optional output folder to save the output report. 
.PARAMETER SmtpServer
	Specifies the optional email server to send the output report. 
.PARAMETER SmtpPort
	Specifies the SMTP port. 
	Default is 25.
.PARAMETER UseSSL
	Specifies whether to use SSL for the SmtpServer.
	Default is False.
.PARAMETER From
	Specifies the username for the From email address.
	If SmtpServer is used, this is a required parameter.
.PARAMETER To
	Specifies the username for the To email address.
	If SmtpServer is used, this is a required parameter.
.PARAMETER Dev
	Clears errors at the beginning of the script.
	Outputs all errors to a text file at the end of the script.
	
	This is used when the script developer requests more troubleshooting data.
	Text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
.PARAMETER ScriptInfo
	Outputs information about the script to a text file.
	Text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
	This parameter has an alias of SI.
.EXAMPLE
	PS C:\PSScript > .\XMS-Report.ps1
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Arnaud Pain" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Arnaud Pain"
	$env:username = Administrator

	Arnaud Pain for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XMS-Report.ps1 -PDF
	
	Will use all default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Arnaud Pain" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Arnaud Pain"
	$env:username = Administrator

	Arnaud Pain for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XMS-Report.ps1 -TEXT

	Will use all default values and save the document as a formatted text file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Arnaud Pain" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Arnaud Pain"
	$env:username = Administrator

	Arnaud Pain for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XMS-Report.ps1 -HTML

	This parameter is reserved for a future update and no output is created at this time.
	
	Will use all default values and save the document as an HTML file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Arnaud Pain" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Arnaud Pain"
	$env:username = Administrator

	Arnaud Pain for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript .\XMS-Report.ps1 -CompanyName "Arnaud Pain Consulting" -CoverPage "Mod" -UserName "Arnaud Pain"

	Will use:
		Arnaud Pain Consulting for the Company Name.
		Mod for the Cover Page format.
		Arnaud Pain for the User Name.
.EXAMPLE
	PS C:\PSScript .\XMS-Report.ps1 -CompanyName "Sherlock Holmes Consulting" `
	-CoverPage Exposure -UserName "Dr. Watson" `
	-CompanyAddress "221B Baker Street, London, England" `
	-CompanyFax "+44 1753 276600" `
	-CompanyPhone "+44 1753 276200"

	Will use:
		Sherlock Holmes Consulting for the Company Name.
		Exposure for the Cover Page format.
		Dr. Watson for the User Name.
		221B Baker Street, London, England for the Company Address.
		+44 1753 276600 for the Company Fax.
		+44 1753 276200 for the Compnay Phone.
.EXAMPLE
	PS C:\PSScript .\XMS-Report.ps1 -CompanyName "Sherlock Holmes Consulting" `
	-CoverPage Facet -UserName "Dr. Watson" `
	-CompanyEmail SuperSleuth@SherlockHolmes.com

	Will use:
		Sherlock Holmes Consulting for the Company Name.
		Facet for the Cover Page format.
		Dr. Watson for the User Name.
		SuperSleuth@SherlockHolmes.com for the Compnay Email.
.EXAMPLE
	PS C:\PSScript .\XMS-Report.ps1 -CN "Arnaud Pain Consulting" -CP "Mod" -UN "Arnaud Pain"

	Will use:
		Arnaud Pain Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Arnaud Pain for the User Name (alias UN).
.EXAMPLE
	PS C:\PSScript > .\XMS-Report.ps1 -AddDateTime
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Arnaud Pain" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Arnaud Pain"
	$env:username = Administrator

	Arnaud Pain for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2017 at 6PM is 2017-06-01_1800.
	Output filename will be Script_Template_2017-06-01_1800.docx
.EXAMPLE
	PS C:\PSScript > .\XMS-Report.ps1 -PDF -AddDateTime
	
	Will use all default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Arnaud Pain" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Arnaud Pain"
	$env:username = Administrator

	Arnaud Pain for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2017 at 6PM is 2017-06-01_1800.
	Output filename will be Script_Template_2017-06-01_1800.PDF

.EXAMPLE
	PS C:\PSScript > .\XMS-Report.ps1 -Folder \\FileServer\ShareName
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Arnaud Pain" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Arnaud Pain"
	$env:username = Administrator

	Arnaud Pain for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Output file will be saved in the path \\FileServer\ShareName
.EXAMPLE
	PS C:\PSScript > .\XMS-Report.ps1 -SmtpServer mail.domain.tld -From XDAdmin@domain.tld -To ITGroup@domain.tld
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Arnaud Pain" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Arnaud Pain"
	$env:username = Administrator

	Arnaud Pain for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Script will use the email server mail.domain.tld, sending from XDAdmin@domain.tld, sending to ITGroup@domain.tld.
	Script will use the default SMTP port 25 and will not use SSL.
	If the current user's credentials are not valid to send email, the user will be prompted to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\XMS-Report.ps1 -SmtpServer smtp.office365.com -SmtpPort 587 -UseSSL -From Arnaud@ArnaudPain.com -To ITGroup@ArnaudPain.com
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Arnaud Pain" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Arnaud Pain"
	$env:username = Administrator

	Arnaud Pain for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Script will use the email server smtp.office365.com on port 587 using SSL, sending from Arnaud@ArnaudPain.com, sending to ITGroup@ArnaudPain.com.
	If the current user's credentials are not valid to send email, the user will be prompted to enter valid credentials.
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  
	This script creates a Word, PDF, Formatted Text or HTML document.
.NOTES
	NAME: XMS-Report.ps1
	VERSION: 1.01
	AUTHOR: Arnaud Pain, Carl Webster, Michael B. Smith, Iain Brighton, Jeff Wouters, Barry Schiffer, Jim Moyle
	LASTEDIT: December 05, 2017
#>

#endregion


#region script parameters
#thanks to @jeffwouters and Michael B. Smith for helping Carl Webster with these parameters
[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "Word") ]

Param(
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Switch]$MSWord=$False,

	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Switch]$PDF=$False,

	[parameter(ParameterSetName="Text",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Switch]$Text=$False,

	[parameter(ParameterSetName="HTML",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Switch]$HTML=$False,

	[parameter(Mandatory=$False)] 
	[Switch]$AddDateTime=$False,
	
	[parameter(Mandatory=$False)] 
	[string]$Folder="",
	
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Alias("CN")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyName="",
    
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Alias("CA")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyAddress="",
    
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Alias("CE")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyEmail="",
    
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Alias("CF")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyFax="",
    
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Alias("CPh")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyPhone="",
    
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Alias("CP")]
	[ValidateNotNullOrEmpty()]
	[string]$CoverPage="Sideline", 

	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Alias("UN")]
	[ValidateNotNullOrEmpty()]
	[string]$UserName=$env:username,

	[parameter(ParameterSetName="SMTP",Mandatory=$True)] 
	[string]$SmtpServer="",

	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[int]$SmtpPort=25,

	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[switch]$UseSSL=$False,

	[parameter(ParameterSetName="SMTP",Mandatory=$True)] 
	[string]$From="",

	[parameter(ParameterSetName="SMTP",Mandatory=$True)] 
	[string]$To="",

	[parameter(Mandatory=$False)] 
	[Switch]$Dev=$False,
	
	[parameter(Mandatory=$False)] 
	[Alias("SI")]
	[Switch]$ScriptInfo=$False
	
	)
#endregion

#region script change log	
#arnaud.pain@arnaud.biz
#@arnaud_pain on Twitter
#http://arnaudpain.com
#Created on September 9, 2017
#endregion

#region initial variable testing and setup
Set-StrictMode -Version 2

#force  on
$PSDefaultParameterValues = @{"*:Verbose"=$True}
$SaveEAPreference = $ErrorActionPreference
$ErrorActionPreference = 'SilentlyContinue'

If($Dev)
{
	$Error.Clear()
	$Script:DevErrorFile = "$($pwd.Path)\XMS-ReportScriptErrors_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
}

If($Null -eq $MSWord)
{
	If($Text -or $HTML -or $PDF)
	{
		$MSWord = $False
	}
	Else
	{
		$MSWord = $True
	}
}

If($MSWord -eq $False -and $PDF -eq $False -and $Text -eq $False -and $HTML -eq $False)
{
	$MSWord = $True
}

Write-Verbose "$(Get-Date): Testing output parameters"

If($MSWord)
{
	Write-Verbose "$(Get-Date): MSWord is set"
}
ElseIf($PDF)
{
	Write-Verbose "$(Get-Date): PDF is set"
}
ElseIf($Text)
{
	Write-Verbose "$(Get-Date): Text is set"
}
ElseIf($HTML)
{
	Write-Verbose "$(Get-Date): HTML is set"
}
Else
{
	$ErrorActionPreference = $SaveEAPreference
	Write-Verbose "$(Get-Date): Unable to determine output parameter"
	If($Null -eq $MSWord)
	{
		Write-Verbose "$(Get-Date): MSWord is Null"
	}
	ElseIf($Null -eq $PDF)
	{
		Write-Verbose "$(Get-Date): PDF is Null"
	}
	ElseIf($Null -eq $Text)
	{
		Write-Verbose "$(Get-Date): Text is Null"
	}
	ElseIf($Null -eq $HTML)
	{
		Write-Verbose "$(Get-Date): HTML is Null"
	}
	Else
	{
		Write-Verbose "$(Get-Date): MSWord is $($MSWord)"
		Write-Verbose "$(Get-Date): PDF is $($PDF)"
		Write-Verbose "$(Get-Date): Text is $($Text)"
		Write-Verbose "$(Get-Date): HTML is $($HTML)"
	}
	Write-Error "Unable to determine output parameter.  Script cannot continue"
	Exit
}

If($Folder -ne "")
{
	Write-Verbose "$(Get-Date): Testing folder path"
	#does it exist
	If(Test-Path $Folder -EA 0)
	{
		#it exists, now check to see if it is a folder and not a file
		If(Test-Path $Folder -pathType Container -EA 0)
		{
			#it exists and it is a folder
			Write-Verbose "$(Get-Date): Folder path $Folder exists and is a folder"
		}
		Else
		{
			#it exists but it is a file not a folder
			Write-Error "Folder $Folder is a file, not a folder.  Script cannot continue"
			Exit
		}
	}
	Else
	{
		#does not exist
		Write-Error "Folder $Folder does not exist.  Script cannot continue"
		Exit
	}
}

#endregion

#region initialize variables for word html and text
[string]$Script:RunningOS = (Get-WmiObject -class Win32_OperatingSystem -EA 0).Caption

If($MSWord -or $PDF)
{
	#try and fix the issue with the $CompanyName variable
	$Script:CoName = $CompanyName
	Write-Verbose "$(Get-Date): CoName is $($Script:CoName)"
	
	#the following values were attained from 
	#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/
	#http://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx
	[int]$wdAlignPageNumberRight = 2
	[long]$wdColorGray15 = 14277081
	[long]$wdColorGray15 = 14277081
	[long]$wdColorRoyalBlue = 13459258
	[int]$wdMove = 0
	[int]$wdSeekMainDocument = 0
	[int]$wdSeekPrimaryFooter = 4
	[int]$wdStory = 6
	[long]$wdColorRed = 255
	[long]$wdColorWhite = 16777215
	[int]$wdColorBlack = 0
	[int]$wdWord2007 = 12
	[int]$wdWord2010 = 14
	[int]$wdWord2013 = 15
	[int]$wdWord2016 = 16
	[int]$wdFormatDocumentDefault = 16
	[int]$wdFormatPDF = 17
	#http://blogs.technet.com/b/heyscriptingguy/archive/2006/03/01/how-can-i-right-align-a-single-column-in-a-word-table.aspx
	#http://msdn.microsoft.com/en-us/library/office/ff835817%28v=office.15%29.aspx
	[int]$wdAlignParagraphLeft = 0
	[int]$wdAlignParagraphCenter = 1
	[int]$wdAlignParagraphRight = 2
	#http://msdn.microsoft.com/en-us/library/office/ff193345%28v=office.15%29.aspx
	[int]$wdCellAlignVerticalTop = 0
	[int]$wdCellAlignVerticalCenter = 1
	[int]$wdCellAlignVerticalBottom = 2
	#http://msdn.microsoft.com/en-us/library/office/ff844856%28v=office.15%29.aspx
	[int]$wdAutoFitFixed = 0
	[int]$wdAutoFitContent = 1
	[int]$wdAutoFitWindow = 2
	#http://msdn.microsoft.com/en-us/library/office/ff821928%28v=office.15%29.aspx
	[int]$wdAdjustNone = 0
	[int]$wdAdjustProportional = 1
	[int]$wdAdjustFirstColumn = 2
	[int]$wdAdjustSameWidth = 3

	[int]$PointsPerTabStop = 36
	[int]$Indent0TabStops = 0 * $PointsPerTabStop
	[int]$Indent1TabStops = 1 * $PointsPerTabStop
	[int]$Indent2TabStops = 2 * $PointsPerTabStop
	[int]$Indent3TabStops = 3 * $PointsPerTabStop
	[int]$Indent4TabStops = 4 * $PointsPerTabStop

	# http://www.thedoctools.com/index.php?show=wt_style_names_english_danish_german_french
	[int]$wdStyleHeading1 = -2
	[int]$wdStyleHeading2 = -3
	[int]$wdStyleHeading3 = -4
	[int]$wdStyleHeading4 = -5
	[int]$wdStyleNoSpacing = -158
	[int]$wdTableGrid = -155

	#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/org/codehaus/groovy/scriptom/tlb/office/word/WdLineStyle.html
	[int]$wdLineStyleNone = 0
	[int]$wdLineStyleSingle = 1

	[int]$wdHeadingFormatTrue = -1
	[int]$wdHeadingFormatFalse = 0 
	
	#portrait and landscape
	$wdOrientLandscape = 1
	$wdOrientPortrait = 0
}

If($HTML)
{
    Set htmlredmask         -Option AllScope -Value "#FF0000" 4>$Null
    Set htmlcyanmask        -Option AllScope -Value "#00FFFF" 4>$Null
    Set htmlbluemask        -Option AllScope -Value "#0000FF" 4>$Null
    Set htmldarkbluemask    -Option AllScope -Value "#0000A0" 4>$Null
    Set htmllightbluemask   -Option AllScope -Value "#ADD8E6" 4>$Null
    Set htmlpurplemask      -Option AllScope -Value "#800080" 4>$Null
    Set htmlyellowmask      -Option AllScope -Value "#FFFF00" 4>$Null
    Set htmllimemask        -Option AllScope -Value "#00FF00" 4>$Null
    Set htmlmagentamask     -Option AllScope -Value "#FF00FF" 4>$Null
    Set htmlwhitemask       -Option AllScope -Value "#FFFFFF" 4>$Null
    Set htmlsilvermask      -Option AllScope -Value "#C0C0C0" 4>$Null
    Set htmlgraymask        -Option AllScope -Value "#808080" 4>$Null
    Set htmlblackmask       -Option AllScope -Value "#000000" 4>$Null
    Set htmlorangemask      -Option AllScope -Value "#FFA500" 4>$Null
    Set htmlmaroonmask      -Option AllScope -Value "#800000" 4>$Null
    Set htmlgreenmask       -Option AllScope -Value "#008000" 4>$Null
    Set htmlolivemask       -Option AllScope -Value "#808000" 4>$Null

    Set htmlbold        -Option AllScope -Value 1 4>$Null
    Set htmlitalics     -Option AllScope -Value 2 4>$Null
    Set htmlred         -Option AllScope -Value 4 4>$Null
    Set htmlcyan        -Option AllScope -Value 8 4>$Null
    Set htmlblue        -Option AllScope -Value 16 4>$Null
    Set htmldarkblue    -Option AllScope -Value 32 4>$Null
    Set htmllightblue   -Option AllScope -Value 64 4>$Null
    Set htmlpurple      -Option AllScope -Value 128 4>$Null
    Set htmlyellow      -Option AllScope -Value 256 4>$Null
    Set htmllime        -Option AllScope -Value 512 4>$Null
    Set htmlmagenta     -Option AllScope -Value 1024 4>$Null
    Set htmlwhite       -Option AllScope -Value 2048 4>$Null
    Set htmlsilver      -Option AllScope -Value 4096 4>$Null
    Set htmlgray        -Option AllScope -Value 8192 4>$Null
    Set htmlolive       -Option AllScope -Value 16384 4>$Null
    Set htmlorange      -Option AllScope -Value 32768 4>$Null
    Set htmlmaroon      -Option AllScope -Value 65536 4>$Null
    Set htmlgreen       -Option AllScope -Value 131072 4>$Null
    Set htmlblack       -Option AllScope -Value 262144 4>$Null
}

If($TEXT)
{
	$global:output = ""
}
#endregion

#region word specific functions
Function SetWordHashTable
{
	Param([string]$CultureCode)

	#optimized by Michael B. Smith
	
	# DE and FR translations for Word 2010 by Vladimir Radojevic
	# Vladimir.Radojevic@Commerzreal.com

	# DA translations for Word 2010 by Thomas Daugaard
	# Citrix Infrastructure Specialist at edgemo A/S

	# CA translations by Javier Sanchez 
	# CEO & Founder 101 Consulting

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish
	#zh - Chinese
	
	[string]$toc = $(
		Switch ($CultureCode)
		{
			'ca-'	{ 'Taula automática 2'; Break }
			'da-'	{ 'Automatisk tabel 2'; Break }
			'de-'	{ 'Automatische Tabelle 2'; Break }
			'en-'	{ 'Automatic Table 2'; Break }
			'es-'	{ 'Tabla automática 2'; Break }
			'fi-'	{ 'Automaattinen taulukko 2'; Break }
			'fr-'	{ 'Table Automatique 2'; Break }
			'nb-'	{ 'Automatisk tabell 2'; Break }
			'nl-'	{ 'Automatische inhoudsopgave 2'; Break }
			'pt-'	{ 'Sumário Automático 2'; Break }
			'sv-'	{ 'Automatisk innehållsförteckning2'; Break }
			'zh-'	{ '自动目录 2'; Break }
		}
	)

	$Script:myHash                      = @{}
	$Script:myHash.Word_TableOfContents = $toc
	$Script:myHash.Word_NoSpacing       = $wdStyleNoSpacing
	$Script:myHash.Word_Heading1        = $wdStyleheading1
	$Script:myHash.Word_Heading2        = $wdStyleheading2
	$Script:myHash.Word_Heading3        = $wdStyleheading3
	$Script:myHash.Word_Heading4        = $wdStyleheading4
	$Script:myHash.Word_TableGrid       = $wdTableGrid
}

Function GetCulture
{
	Param([int]$WordValue)
	
	#codes obtained from http://support.microsoft.com/kb/221435
	#http://msdn.microsoft.com/en-us/library/bb213877(v=office.12).aspx
	$CatalanArray = 1027
	$ChineseArray = 2052,3076,5124,4100
	$DanishArray = 1030
	$DutchArray = 2067, 1043
	$EnglishArray = 3081, 10249, 4105, 9225, 6153, 8201, 5129, 13321, 7177, 11273, 2057, 1033, 12297
	$FinnishArray = 1035
	$FrenchArray = 2060, 1036, 11276, 3084, 12300, 5132, 13324, 6156, 8204, 10252, 7180, 9228, 4108
	$GermanArray = 1031, 3079, 5127, 4103, 2055
	$NorwegianArray = 1044, 2068
	$PortugueseArray = 1046, 2070
	$SpanishArray = 1034, 11274, 16394, 13322, 9226, 5130, 7178, 12298, 17418, 4106, 18442, 19466, 6154, 15370, 10250, 20490, 3082, 14346, 8202
	$SwedishArray = 1053, 2077

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish
	#zh - Chinese

	Switch ($WordValue)
	{
		{$CatalanArray -contains $_} {$CultureCode = "ca-"}
		{$ChineseArray -contains $_} {$CultureCode = "zh-"}
		{$DanishArray -contains $_} {$CultureCode = "da-"}
		{$DutchArray -contains $_} {$CultureCode = "nl-"}
		{$EnglishArray -contains $_} {$CultureCode = "en-"}
		{$FinnishArray -contains $_} {$CultureCode = "fi-"}
		{$FrenchArray -contains $_} {$CultureCode = "fr-"}
		{$GermanArray -contains $_} {$CultureCode = "de-"}
		{$NorwegianArray -contains $_} {$CultureCode = "nb-"}
		{$PortugueseArray -contains $_} {$CultureCode = "pt-"}
		{$SpanishArray -contains $_} {$CultureCode = "es-"}
		{$SwedishArray -contains $_} {$CultureCode = "sv-"}
		Default {$CultureCode = "en-"}
	}
	
	Return $CultureCode
}

Function ValidateCoverPage
{
	Param([int]$xWordVersion, [string]$xCP, [string]$CultureCode)
	
	$xArray = ""
	
	Switch ($CultureCode)
	{
		'ca-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització principal", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Anual", "Austin", "Conservador",
					"Contrast", "Cubicles", "Diplomàtic", "Exposició",
					"Línia lateral", "Mod", "Mosiac", "Moviment", "Paper de diari",
					"Perspectiva", "Piles", "Quadrícula", "Sobri",
					"Transcendir", "Trencaclosques")
				}
			}

		'da-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "BevægElse", "Brusen", "Facet", "Filigran", 
					"Gitter", "Integral", "Ion (lys)", "Ion (mørk)", 
					"Retro", "Semafor", "Sidelinje", "Stribet", 
					"Udsnit (lys)", "Udsnit (mørk)", "Visningsmaster")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("BevægElse", "Brusen", "Ion (lys)", "Filigran",
					"Retro", "Semafor", "Visningsmaster", "Integral",
					"Facet", "Gitter", "Stribet", "Sidelinje", "Udsnit (lys)",
					"Udsnit (mørk)", "Ion (mørk)", "Austin")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("BevægElse", "Moderat", "Perspektiv", "Firkanter",
					"Overskrid", "Alfabet", "Kontrast", "Stakke", "Fliser", "Gåde",
					"Gitter", "Austin", "Eksponering", "Sidelinje", "Enkel",
					"Nålestribet", "Årlig", "Avispapir", "Tradionel")
				}
			}

		'de-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Bewegung", "Facette", "Filigran", 
					"Gebändert", "Integral", "Ion (dunkel)", "Ion (hell)", 
					"Pfiff", "Randlinie", "Raster", "Rückblick", 
					"Segment (dunkel)", "Segment (hell)", "Semaphor", 
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Semaphor", "Segment (hell)", "Ion (hell)",
					"Raster", "Ion (dunkel)", "Filigran", "Rückblick", "Pfiff",
					"ViewMaster", "Segment (dunkel)", "Verbunden", "Bewegung",
					"Randlinie", "Austin", "Integral", "Facette")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Austin", "Bewegung", "Durchscheinend",
					"Herausgestellt", "Jährlich", "Kacheln", "Kontrast", "Kubistisch",
					"Modern", "Nadelstreifen", "Perspektive", "Puzzle", "Randlinie",
					"Raster", "Schlicht", "Stapel", "Traditionell", "Zeitungspapier")
				}
			}

		'en-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
					"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
					"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
					"Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
					"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
					"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
				}
			}

		'es-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Con bandas", "Cortar (oscuro)", "Cuadrícula", 
					"Whisp", "Faceta", "Filigrana", "Integral", "Ion (claro)", 
					"Ion (oscuro)", "Línea lateral", "Movimiento", "Retrospectiva", 
					"Semáforo", "Slice (luz)", "Vista principal", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Whisp", "Vista principal", "Filigrana", "Austin",
					"Slice (luz)", "Faceta", "Semáforo", "Retrospectiva", "Cuadrícula",
					"Movimiento", "Cortar (oscuro)", "Línea lateral", "Ion (oscuro)",
					"Ion (claro)", "Integral", "Con bandas")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Anual", "Austero", "Austin", "Conservador",
					"Contraste", "Cuadrícula", "Cubículos", "Exposición", "Línea lateral",
					"Moderno", "Mosaicos", "Movimiento", "Papel periódico",
					"Perspectiva", "Pilas", "Puzzle", "Rayas", "Sobrepasar")
				}
			}

		'fi-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
					"Kuiskaus", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
					"Kiehkura", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aakkoset", "Askeettinen", "Austin", "Kontrasti",
					"Laatikot", "Liike", "Liituraita", "Mod", "Osittain peitossa",
					"Palapeli", "Perinteinen", "Perspektiivi", "Pinot", "Ruudukko",
					"Ruudut", "Sanomalehtipaperi", "Sivussa", "Vuotuinen", "Ylitys")
				}
			}

		'fr-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("À bandes", "Austin", "Facette", "Filigrane", 
					"Guide", "Intégrale", "Ion (clair)", "Ion (foncé)", 
					"Lignes latérales", "Quadrillage", "Rétrospective", "Secteur (clair)", 
					"Secteur (foncé)", "Sémaphore", "ViewMaster", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annuel", "Austère", "Austin", 
					"Blocs empilés", "Classique", "Contraste", "Emplacements de bureau", 
					"Exposition", "Guide", "Ligne latérale", "Moderne", 
					"Mosaïques", "Mots croisés", "Papier journal", "Perspective",
					"Quadrillage", "Rayures fines", "Transcendant")
				}
			}

		'nb-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "BevegElse", "Dempet", "Fasett", "Filigran",
					"Integral", "Ion (lys)", "Ion (mørk)", "Retrospekt", "Rutenett",
					"Sektor (lys)", "Sektor (mørk)", "Semafor", "Sidelinje", "Stripet",
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Årlig", "Avistrykk", "Austin", "Avlukker",
					"BevegElse", "Engasjement", "Enkel", "Fliser", "Konservativ",
					"Kontrast", "Mod", "Perspektiv", "Puslespill", "Rutenett", "Sidelinje",
					"Smale striper", "Stabler", "Transcenderende")
				}
			}

		'nl-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Beweging", "Facet", "Filigraan", "Gestreept",
					"Integraal", "Ion (donker)", "Ion (licht)", "Raster",
					"Segment (Light)", "Semafoor", "Slice (donker)", "Spriet",
					"Terugblik", "Terzijde", "ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aantrekkelijk", "Alfabet", "Austin", "Bescheiden",
					"Beweging", "Blikvanger", "Contrast", "Eenvoudig", "Jaarlijks",
					"Krantenpapier", "Krijtstreep", "Kubussen", "Mod", "Perspectief",
					"Puzzel", "Raster", "Stapels",
					"Tegels", "Terzijde")
				}
			}

		'pt-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Animação", "Austin", "Em Tiras", "Exibição Mestra",
					"Faceta", "Fatia (Clara)", "Fatia (Escura)", "Filete", "Filigrana", 
					"Grade", "Integral", "Íon (Claro)", "Íon (Escuro)", "Linha Lateral",
					"Retrospectiva", "Semáforo")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Animação", "Anual", "Austero", "Austin", "Baias",
					"Conservador", "Contraste", "Exposição", "Grade", "Ladrilhos",
					"Linha Lateral", "Listras", "Mod", "Papel Jornal", "Perspectiva", "Pilhas",
					"Quebra-cabeça", "Transcend")
				}
			}

		'sv-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Band", "Fasett", "Filigran", "Integrerad", "Jon (ljust)",
					"Jon (mörkt)", "Knippe", "Rutnät", "RörElse", "Sektor (ljus)", "Sektor (mörk)",
					"Semafor", "Sidlinje", "VisaHuvudsida", "Återblick")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabetmönster", "Austin", "Enkelt", "Exponering", "Konservativt",
					"Kontrast", "Kritstreck", "Kuber", "Perspektiv", "Plattor", "Pussel", "Rutnät",
					"RörElse", "Sidlinje", "Sobert", "Staplat", "Tidningspapper", "Årligt",
					"Övergående")
				}
			}

		'zh-'	{
				If($xWordVersion -eq $wdWord2010 -or $xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("奥斯汀", "边线型", "花丝", "怀旧", "积分",
					"离子(浅色)", "离子(深色)", "母版型", "平面", "切片(浅色)",
					"切片(深色)", "丝状", "网格", "镶边", "信号灯",
					"运动型")
				}
			}

		Default	{
					If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
					{
						$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
						"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
						"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
						"Whisp")
					}
					ElseIf($xWordVersion -eq $wdWord2010)
					{
						$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
						"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
						"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
					}
				}
	}
	
	If($xArray -contains $xCP)
	{
		$xArray = $Null
		Return $True
	}
	Else
	{
		$xArray = $Null
		Return $False
	}
}

Function CheckWordPrereq
{
	If((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Word.Application) -eq $False)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Host "`n`n`t`tThis script directly outputs to Microsoft Word, please install Microsoft Word`n`n"
		Exit
	}

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId
	
	#Find out if winword is running in our session
	[bool]$wordrunning = ((Get-Process 'WinWord' -ea 0)|?{$_.SessionId -eq $SessionID}) -ne $Null
	If($wordrunning)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Host "`n`n`tPlease close all instances of Microsoft Word before running this report.`n`n"
		Exit
	}
}

Function ValidateCompanyName
{
	[bool]$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	If($xResult)
	{
		Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	}
	Else
	{
		$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		If($xResult)
		{
			Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		}
		Else
		{
			Return ""
		}
	}
}

Function Set-DocumentProperty 
{
    <#
	.SYNOPSIS
	Function to set the Title Page document properties in MS Word
	.DESCRIPTION
	Long description
	.PARAMETER Document
	Current Document Object
	.PARAMETER DocProperty
	Parameter description
	.PARAMETER Value
	Parameter description
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Title -Value 'MyTitle'
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Company -Value 'MyCompany'
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Author -Value 'Jim Moyle'
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Subject -Value 'MySubjectTitle'
	.NOTES
	Function Created by Jim Moyle June 2017
	Twitter : @JimMoyle
	#>
    param (
        [object]$Document,
        [String]$DocProperty,
        [string]$Value
    )
    try {
        $binding = "System.Reflection.BindingFlags" -as [type]
        $builtInProperties = $Document.BuiltInDocumentProperties
        $property = [System.__ComObject].invokemember("item", $binding::GetProperty, $null, $BuiltinProperties, $DocProperty)
        [System.__ComObject].invokemember("value", $binding::SetProperty, $null, $property, $Value)
    }
    catch {
        Write-Warning "Failed to set $DocProperty to $Value"
    }
}

Function FindWordDocumentEnd
{
	#return focus to main document    
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument
	#move to the end of the current document
	$Script:Selection.EndKey($wdStory,$wdMove) | Out-Null
}

Function SetupWord
{
	Write-Verbose "$(Get-Date): Setting up Word"
    
	# Setup word for output
	Write-Verbose "$(Get-Date): Create Word comObject."
	$Script:Word = New-Object -comobject "Word.Application" -EA 0 4>$Null
	
	If(!$? -or $Null -eq $Script:Word)
	{
		Write-Warning "The Word object could not be created.  You may need to repair your Word installation."
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tThe Word object could not be created.  You may need to repair your Word installation.`n`n`t`tScript cannot continue.`n`n"
		Exit
	}

	Write-Verbose "$(Get-Date): Determine Word language value"
	If( ( validStateProp $Script:Word Language Value__ ) )
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language.Value__
	}
	Else
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language
	}

	If(!($Script:WordLanguageValue -gt -1))
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tUnable to determine the Word language value.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}
	Write-Verbose "$(Get-Date): Word language value is $($Script:WordLanguageValue)"
	
	$Script:WordCultureCode = GetCulture $Script:WordLanguageValue
	
	SetWordHashTable $Script:WordCultureCode
	
	[int]$Script:WordVersion = [int]$Script:Word.Version
	If($Script:WordVersion -eq $wdWord2016)
	{
		$Script:WordProduct = "Word 2016"
	}
	ElseIf($Script:WordVersion -eq $wdWord2013)
	{
		$Script:WordProduct = "Word 2013"
	}
	ElseIf($Script:WordVersion -eq $wdWord2010)
	{
		$Script:WordProduct = "Word 2010"
	}
	ElseIf($Script:WordVersion -eq $wdWord2007)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tMicrosoft Word 2007 is no longer supported.`n`n`t`tScript will end.`n`n"
		AbortScript
	}
	Else
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tYou are running an untested or unsupported version of Microsoft Word.`n`n`t`tScript will end.`n`n`t`tPlease send info on your version of Word to webster@carlwebster.com`n`n"
		AbortScript
	}

	#only validate CompanyName if the field is blank
	If([String]::IsNullOrEmpty($Script:CoName))
	{
		Write-Verbose "$(Get-Date): Company name is blank.  Retrieve company name from registry."
		$TmpName = ValidateCompanyName
		
		If([String]::IsNullOrEmpty($TmpName))
		{
			Write-Warning "`n`n`t`tCompany Name is blank so Cover Page will not show a Company Name."
			Write-Warning "`n`t`tCheck HKCU:\Software\Microsoft\Office\Common\UserInfo for Company or CompanyName value."
			Write-Warning "`n`t`tYou may want to use the -CompanyName parameter if you need a Company Name on the cover page.`n`n"
		}
		Else
		{
			$Script:CoName = $TmpName
			Write-Verbose "$(Get-Date): Updated company name to $($Script:CoName)"
		}
	}

	If($Script:WordCultureCode -ne "en-")
	{
		Write-Verbose "$(Get-Date): Check Default Cover Page for $($WordCultureCode)"
		[bool]$CPChanged = $False
		Switch ($Script:WordCultureCode)
		{
			'ca-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Línia lateral"
						$CPChanged = $True
					}
				}

			'da-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidelinje"
						$CPChanged = $True
					}
				}

			'de-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Randlinie"
						$CPChanged = $True
					}
				}

			'es-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Línea lateral"
						$CPChanged = $True
					}
				}

			'fi-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sivussa"
						$CPChanged = $True
					}
				}

			'fr-'	{
					If($CoverPage -eq "Sideline")
					{
						If($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016)
						{
							$CoverPage = "Lignes latérales"
							$CPChanged = $True
						}
						Else
						{
							$CoverPage = "Ligne latérale"
							$CPChanged = $True
						}
					}
				}

			'nb-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidelinje"
						$CPChanged = $True
					}
				}

			'nl-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Terzijde"
						$CPChanged = $True
					}
				}

			'pt-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Linha Lateral"
						$CPChanged = $True
					}
				}

			'sv-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidlinje"
						$CPChanged = $True
					}
				}

			'zh-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "边线型"
						$CPChanged = $True
					}
				}
		}

		If($CPChanged)
		{
			Write-Verbose "$(Get-Date): Changed Default Cover Page from Sideline to $($CoverPage)"
		}
	}

	Write-Verbose "$(Get-Date): Validate cover page $($CoverPage) for culture code $($Script:WordCultureCode)"
	[bool]$ValidCP = $False
	
	$ValidCP = ValidateCoverPage $Script:WordVersion $CoverPage $Script:WordCultureCode
	
	If(!$ValidCP)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Verbose "$(Get-Date): Word language value $($Script:WordLanguageValue)"
		Write-Verbose "$(Get-Date): Culture code $($Script:WordCultureCode)"
		Write-Error "`n`n`t`tFor $($Script:WordProduct), $($CoverPage) is not a valid Cover Page option.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}

	ShowScriptOptions

	$Script:Word.Visible = $False

	#http://jdhitsolutions.com/blog/2012/05/san-diego-2012-powershell-deep-dive-slides-and-demos/
	#using Jeff's Demo-WordReport.ps1 file for examples
	Write-Verbose "$(Get-Date): Load Word Templates"

	[bool]$Script:CoverPagesExist = $False
	[bool]$BuildingBlocksExist = $False

	$Script:Word.Templates.LoadBuildingBlocks()
	#word 2010/2013/2016
	$BuildingBlocksCollection = $Script:Word.Templates | Where {$_.name -eq "Built-In Building Blocks.dotx"}

	Write-Verbose "$(Get-Date): Attempt to load cover page $($CoverPage)"
	$part = $Null

	$BuildingBlocksCollection | 
	ForEach{
		If ($_.BuildingBlockEntries.Item($CoverPage).Name -eq $CoverPage) 
		{
			$BuildingBlocks = $_
		}
	}        

	If($Null -ne $BuildingBlocks)
	{
		$BuildingBlocksExist = $True

		Try 
		{
			$part = $BuildingBlocks.BuildingBlockEntries.Item($CoverPage)
		}

		Catch
		{
			$part = $Null
		}

		If($Null -ne $part)
		{
			$Script:CoverPagesExist = $True
		}
	}

	If(!$Script:CoverPagesExist)
	{
		Write-Verbose "$(Get-Date): Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
		Write-Warning "Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
		Write-Warning "This report will not have a Cover Page."
	}

	Write-Verbose "$(Get-Date): Create empty word doc"
	$Script:Doc = $Script:Word.Documents.Add()
	If($Null -eq $Script:Doc)
	{
		Write-Verbose "$(Get-Date): "
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tAn empty Word document could not be created.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}

	$Script:Selection = $Script:Word.Selection
	If($Null -eq $Script:Selection)
	{
		Write-Verbose "$(Get-Date): "
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tAn unknown error happened selecting the entire Word document for default formatting options.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}

	#set Default tab stops to 1/2 inch (this line is not from Jeff Hicks)
	#36 = .50"
	$Script:Word.ActiveDocument.DefaultTabStop = 36

	#Disable Spell and Grammar Check to resolve issue and improve performance (from Pat Coughlin)
	Write-Verbose "$(Get-Date): Disable grammar and spell checking"
	#bug reported 1-Apr-2014 by Tim Mangan
	#save current options first before turning them off
	$Script:CurrentGrammarOption = $Script:Word.Options.CheckGrammarAsYouType
	$Script:CurrentSpellingOption = $Script:Word.Options.CheckSpellingAsYouType
	$Script:Word.Options.CheckGrammarAsYouType = $False
	$Script:Word.Options.CheckSpellingAsYouType = $False

	If($BuildingBlocksExist)
	{
		#insert new page, getting ready for table of contents
		Write-Verbose "$(Get-Date): Insert new page, getting ready for table of contents"
		$part.Insert($Script:Selection.Range,$True) | Out-Null
		$Script:Selection.InsertNewPage()

		#table of contents
		Write-Verbose "$(Get-Date): Table of Contents - $($Script:MyHash.Word_TableOfContents)"
		$toc = $BuildingBlocks.BuildingBlockEntries.Item($Script:MyHash.Word_TableOfContents)
		If($Null -eq $toc)
		{
			Write-Verbose "$(Get-Date): "
			Write-Verbose "$(Get-Date): Table of Content - $($Script:MyHash.Word_TableOfContents) could not be retrieved."
			Write-Warning "This report will not have a Table of Contents."
		}
		Else
		{
			$toc.insert($Script:Selection.Range,$True) | Out-Null
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): Table of Contents are not installed."
		Write-Warning "Table of Contents are not installed so this report will not have a Table of Contents."
	}

	#set the footer
	Write-Verbose "$(Get-Date): Set the footer"
	[string]$footertext = "Report created by $username"

	#get the footer
	Write-Verbose "$(Get-Date): Get the footer and format font"
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekPrimaryFooter
	#get the footer and format font
	$footers = $Script:Doc.Sections.Last.Footers
	ForEach ($footer in $footers) 
	{
		If($footer.exists) 
		{
			$footer.range.Font.name = "Calibri"
			$footer.range.Font.size = 8
			$footer.range.Font.Italic = $True
			$footer.range.Font.Bold = $True
		}
	} #end ForEach
	Write-Verbose "$(Get-Date): Footer text"
	$Script:Selection.HeaderFooter.Range.Text = $footerText

	#add page numbering
	Write-Verbose "$(Get-Date): Add page numbering"
	$Script:Selection.HeaderFooter.PageNumbers.Add($wdAlignPageNumberRight) | Out-Null

	FindWordDocumentEnd
	Write-Verbose "$(Get-Date):"
	#end of Jeff Hicks 
}

Function UpdateDocumentProperties
{
	Param([string]$AbstractTitle, [string]$SubjectTitle)
	#Update document properties
	If($MSWORD -or $PDF)
	{
		If($Script:CoverPagesExist)
		{
			Write-Verbose "$(Get-Date): Set Cover Page Properties"
			#8-Jun-2017 put these 4 items in alpha order
            Set-DocumentProperty -Document $Script:Doc -DocProperty Author -Value $UserName
            Set-DocumentProperty -Document $Script:Doc -DocProperty Company -Value $Script:CoName
            Set-DocumentProperty -Document $Script:Doc -DocProperty Subject -Value $SubjectTitle
            Set-DocumentProperty -Document $Script:Doc -DocProperty Title -Value $Script:title

			#Get the Coverpage XML part
			$cp = $Script:Doc.CustomXMLParts | Where {$_.NamespaceURI -match "coverPageProps$"}

			#get the abstract XML part
			$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "Abstract"}
			#set the text
			If([String]::IsNullOrEmpty($Script:CoName))
			{
				[string]$abstract = $AbstractTitle
			}
			Else
			{
				[string]$abstract = "$($AbstractTitle) for $($Script:CoName)"
			}
			$ab.Text = $abstract
			
			$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "CompanyAddress"}
			#set the text
			[string]$abstract = $CompanyAddress
			$ab.Text = $abstract

			$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "CompanyEmail"}
			#set the text
			[string]$abstract = $CompanyEmail
			$ab.Text = $abstract

			$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "CompanyFax"}
			#set the text
			[string]$abstract = $CompanyFax
			$ab.Text = $abstract

			$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "CompanyPhone"}
			#set the text
			[string]$abstract = $CompanyPhone
			$ab.Text = $abstract

			$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "PublishDate"}
			#set the text
			[string]$abstract = (Get-Date -Format d).ToString()
			$ab.Text = $abstract

			Write-Verbose "$(Get-Date): Update the Table of Contents"
			#update the Table of Contents
			$Script:Doc.TablesOfContents.item(1).Update()
			$cp = $Null
			$ab = $Null
			$abstract = $Null
		}
	}
}
#endregion

#region registry functions
#http://stackoverflow.com/questions/5648931/test-if-registry-value-exists
# This Function just gets $True or $False
Function Test-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	$key -and $Null -ne $key.GetValue($name, $Null)
}

# Gets the specified registry value or $Null if it is missing
Function Get-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	If($key)
	{
		$key.GetValue($name, $Null)
	}
	Else
	{
		$Null
	}
}
#endregion

#region word, text and html line output functions
Function line
#function created by Michael B. Smith, Exchange MVP
#@essentialexchange on Twitter
#http://TheEssentialExchange.com
#for creating the formatted text report
{
	Param( [int]$tabs = 0, [string]$name = '', [string]$value = '', [string]$newline = "`r`n", [switch]$nonewline )
	While( $tabs -gt 0 ) { $Global:Output += "`t"; $tabs--; }
	If( $nonewline )
	{
		$Global:Output += $name + $value
	}
	Else
	{
		$Global:Output += $name + $value + $newline
	}
}
	
Function WriteWordLine
#Function created by Ryan Revord
#@rsrevord on Twitter
#Function created to make output to Word easy in this script
{
	Param([int]$style=0, 
	[int]$tabs = 0, 
	[string]$name = '', 
	[string]$value = '', 
	[string]$fontName=$Null,
	[int]$fontSize=0,
	[bool]$italics=$False,
	[bool]$boldface=$False,
	[Switch]$nonewline)
	
	#Build output style
	[string]$output = ""
	Switch ($style)
	{
		0 {$Script:Selection.Style = $Script:MyHash.Word_NoSpacing}
		1 {$Script:Selection.Style = $Script:MyHash.Word_Heading1}
		2 {$Script:Selection.Style = $Script:MyHash.Word_Heading2}
		3 {$Script:Selection.Style = $Script:MyHash.Word_Heading3}
		4 {$Script:Selection.Style = $Script:MyHash.Word_Heading4}
		Default {$Script:Selection.Style = $Script:MyHash.Word_NoSpacing}
	}
	
	#build # of tabs
	While($tabs -gt 0)
	{ 
		$output += "`t"; $tabs--; 
	}
 
	If(![String]::IsNullOrEmpty($fontName)) 
	{
		$Script:Selection.Font.name = $fontName
	} 

	If($fontSize -ne 0) 
	{
		$Script:Selection.Font.size = $fontSize
	} 
 
	If($italics -eq $True) 
	{
		$Script:Selection.Font.Italic = $True
	} 
 
	If($boldface -eq $True) 
	{
		$Script:Selection.Font.Bold = $True
	} 

	#output the rest of the parameters.
	$output += $name + $value
	$Script:Selection.TypeText($output)
 
	#test for new WriteWordLine 0.
	If($nonewline)
	{
		# Do nothing.
	} 
	Else 
	{
		$Script:Selection.TypeParagraph()
	}
}

#***********************************************************************************************************
# WriteHTMLLine
#***********************************************************************************************************

<#
.Synopsis
	Writes a line of output for HTML output
.DESCRIPTION
	This function formats an HTML line
.USAGE
	WriteHTMLLine <Style> <Tabs> <Name> <Value> <Font Name> <Font Size> <Options>

	0 for Font Size denotes using the default font size of 2 or 10 point

.EXAMPLE
	WriteHTMLLine 0 0 ""

	Writes a blank line with no style or tab stops, obviously none needed.

.EXAMPLE
	WriteHTMLLine 0 1 "This is a regular line of text indented 1 tab stops"

	Writes a line with 1 tab stop.

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in italics" "" $Null 0 $htmlitalics

	Writes a line omitting font and font size and setting the italics attribute

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in bold" "" $Null 0 $htmlbold

	Writes a line omitting font and font size and setting the bold attribute

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in bold italics" "" $Null 0 ($htmlbold -bor $htmlitalics)

	Writes a line omitting font and font size and setting both italics and bold options

.EXAMPLE	
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in 10 point" "" $Null 2  # 10 point font

	Writes a line using 10 point font

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in Courier New font" "" "Courier New" 0 

	Writes a line using Courier New Font and 0 font point size (default = 2 if set to 0)

.EXAMPLE	
	WriteHTMLLine 0 0 "This is a regular line of RED text indented 0 tab stops with the computer name as data in 10 point Courier New bold italics: " $env:computername "Courier New" 2 ($htmlbold -bor $htmlred -bor $htmlitalics)

	Writes a line using Courier New Font with first and second string values to be used, also uses 10 point font with bold, italics and red color options set.

.NOTES

	Font Size - Unlike word, there is a limited set of font sizes that can be used in HTML.  They are:
		0 - default which actually gives it a 2 or 10 point.
		1 - 7.5 point font size
		2 - 10 point
		3 - 13.5 point
		4 - 15 point
		5 - 18 point
		6 - 24 point
		7 - 36 point
	Any number larger than 7 defaults to 7

	Style - Refers to the headers that are used with output and resemble the headers in word, 
	HTML supports headers h1-h6 and h1-h4 are more commonly used.  Unlike word, H1 will not 
	give you a blue colored font, you will have to set that yourself.

	Colors and Bold/Italics Flags are:

		htmlbold       
		htmlitalics    
		htmlred        
		htmlcyan        
		htmlblue       
		htmldarkblue   
		htmllightblue   
		htmlpurple      
		htmlyellow      
		htmllime       
		htmlmagenta     
		htmlwhite       
		htmlsilver      
		htmlgray       
		htmlolive       
		htmlorange      
		htmlmaroon      
		htmlgreen       
		htmlblack       
#>

Function WriteHTMLLine
#Function created by Ken Avram
#Function created to make output to HTML easy in this script
#headings fixed 12-Oct-2016 by Webster
{
	Param([int]$style=0, 
	[int]$tabs = 0, 
	[string]$name = '', 
	[string]$value = '', 
	[string]$fontName="Calibri",
	[int]$fontSize=1,
	[int]$options=$htmlblack)


	#Build output style
	[string]$output = ""

	If([String]::IsNullOrEmpty($Name))	
	{
		$HTMLBody = "<p></p>"
	}
	Else
	{
		$color = CheckHTMLColor $options

		#build # of tabs

		While($tabs -gt 0)
		{ 
			$output += "&nbsp;&nbsp;&nbsp;&nbsp;"; $tabs--; 
		}

		$HTMLFontName = $fontName		

		$HTMLBody = ""

		If($options -band $htmlitalics) 
		{
			$HTMLBody += "<i>"
		} 

		If($options -band $htmlbold) 
		{
			$HTMLBody += "<b>"
		} 

		#output the rest of the parameters.
		$output += $name + $value

		#added by webster 12-oct-2016
		#if a heading, don't add the <br>
		If($HTMLStyle -eq "")
		{
			$HTMLBody += "<br><font face='" + $HTMLFontName + "' " + "color='" + $color + "' size='"  + $fontsize + "'>"
		}
		Else
		{
			$HTMLBody += "<font face='" + $HTMLFontName + "' " + "color='" + $color + "' size='"  + $fontsize + "'>"
		}
		
		Switch ($style)
		{
			1 {$HTMLStyle = "<h1>"; Break}
			2 {$HTMLStyle = "<h2>"; Break}
			3 {$HTMLStyle = "<h3>"; Break}
			4 {$HTMLStyle = "<h4>"; Break}
			Default {$HTMLStyle = ""; Break}
		}

		$HTMLBody += $HTMLStyle + $output

		Switch ($style)
		{
			1 {$HTMLStyle = "</h1>"; Break}
			2 {$HTMLStyle = "</h2>"; Break}
			3 {$HTMLStyle = "</h3>"; Break}
			4 {$HTMLStyle = "</h4>"; Break}
			Default {$HTMLStyle = ""; Break}
		}

		$HTMLBody += $HTMLStyle +  "</font>"

		If($options -band $htmlitalics) 
		{
			$HTMLBody += "</i>"
		} 

		If($options -band $htmlbold) 
		{
			$HTMLBody += "</b>"
		} 
	}
	
	#added by webster 12-oct-2016
	#if a heading, don't add the <br />
	If($HTMLStyle -eq "")
	{
		$HTMLBody += "<br />"
	}

	out-file -FilePath $Script:FileName1 -Append -InputObject $HTMLBody 4>$Null
}
#endregion

#region HTML table functions
#***********************************************************************************************************
# AddHTMLTable - Called from FormatHTMLTable function
# Created by Ken Avram
# modified by Jake Rutski
#***********************************************************************************************************
Function AddHTMLTable
{
	Param([string]$fontName="Calibri",
	[int]$fontSize=2,
	[int]$colCount=0,
	[int]$rowCount=0,
	[object[]]$rowInfo=@(),
	[object[]]$fixedInfo=@())

	For($rowidx = $RowIndex;$rowidx -le $rowCount;$rowidx++)
	{
		$rd = @($rowInfo[$rowidx - 2])
		$htmlbody = $htmlbody + "<tr>"
		For($columnIndex = 0; $columnIndex -lt $colCount; $columnindex+=2)
		{
			$fontitalics = $False
			$fontbold = $false
			$tmp = CheckHTMLColor $rd[$columnIndex+1]

			If($fixedInfo.Length -eq 0)
			{
				$htmlbody += "<td style=""background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}
			Else
			{
				$htmlbody += "<td style=""width:$($fixedInfo[$columnIndex/2]); background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}

			If($rd[$columnIndex+1] -band $htmlbold)
			{
				$htmlbody += "<b>"
			}
			If($rd[$columnIndex+1] -band $htmlitalics)
			{
				$htmlbody += "<i>"
			}
			If($Null -ne $rd[$columnIndex])
			{
				$cell = $rd[$columnIndex].tostring()
				If($cell -eq " " -or $cell.length -eq 0)
				{
					$htmlbody += "&nbsp;&nbsp;&nbsp;"
				}
				Else
				{
					For($i=0;$i -lt $cell.length;$i++)
					{
						If($cell[$i] -eq " ")
						{
							$htmlbody += "&nbsp;"
						}
						If($cell[$i] -ne " ")
						{
							Break
						}
					}
					$htmlbody += $cell
				}
			}
			Else
			{
				$htmlbody += "&nbsp;&nbsp;&nbsp;"
			}
			If($rd[$columnIndex+1] -band $htmlbold)
			{
				$htmlbody += "</b>"
			}
			If($rd[$columnIndex+1] -band $htmlitalics)
			{
				$htmlbody += "</i>"
			}
			$htmlbody += "</font></td>"
		}
		$htmlbody += "</tr>"
	}
	out-file -FilePath $Script:FileName1 -Append -InputObject $HTMLBody 4>$Null 
}

#***********************************************************************************************************
# FormatHTMLTable 
# Created by Ken Avram
# modified by Jake Rutski
#***********************************************************************************************************

<#
.Synopsis
	Format table for HTML output document
.DESCRIPTION
	This function formats a table for HTML from an array of strings
.PARAMETER noBorder
	If set to $true, a table will be generated without a border (border='0')
.PARAMETER noHeadCols
	This parameter should be used when generating tables without column headers
	Set this parameter equal to the number of columns in the table
.PARAMETER rowArray
	This parameter contains the row data array for the table
.PARAMETER columnArray
	This parameter contains column header data for the table
.PARAMETER fixedWidth
	This parameter contains widths for columns in pixel format ("100px") to override auto column widths
	The variable should contain a width for each column you wish to override the auto-size setting
	For example: $columnWidths = @("100px","110px","120px","130px","140px")

.USAGE
	FormatHTMLTable <Table Header> <Table Format> <Font Name> <Font Size>

.EXAMPLE
	FormatHTMLTable "Table Heading" "auto" "Calibri" 3

	This example formats a table and writes it out into an html file.  All of the parameters are optional
	defaults are used if not supplied.

	for <Table format>, the default is auto which will autofit the text into the columns and adjust to the longest text in that column.  You can also use percentage i.e. 25%
	which will take only 25% of the line and will auto word wrap the text to the next line in the column.  Also, instead of using a percentage, you can use pixels i.e. 400px.

	FormatHTMLTable "Table Heading" "auto" -rowArray $rowData -columnArray $columnData

	This example creates an HTML table with a heading of 'Table Heading', auto column spacing, column header data from $columnData and row data from $rowData

	FormatHTMLTable "Table Heading" -rowArray $rowData -noHeadCols 3

	This example creates an HTML table with a heading of 'Table Heading', auto column spacing, no header, and row data from $rowData

	FormatHTMLTable "Table Heading" -rowArray $rowData -fixedWidth $fixedColumns

	This example creates an HTML table with a heading of 'Table Heading, no header, row data from $rowData, and fixed columns defined by $fixedColumns

.NOTES
	In order to use the formatted table it first has to be loaded with data.  Examples below will show how to load the table:

	First, initialize the table array

	$rowdata = @()

	Then Load the array.  If you are using column headers then load those into the column headers array, otherwise the first line of the table goes into the column headers array
	and the second and subsequent lines go into the $rowdata table as shown below:

	$columnHeaders = @('Display Name',($htmlsilver -bor $htmlbold),'Description',($htmlsilver -bor $htmlbold),'Key',($htmlsilver -bor $htmlbold))

	The first column is the actual name to display, the second are the attributes of the column i.e. color anded with bold or italics.  For the anding, parens are required or it will
	not format correctly.

	This is following by adding rowdata as shown below.  As more columns are added the columns will auto adjust to fit the size of the page.

	$rowdata = @()
	$columnHeaders = @("User Name",($htmlsilver -bor $htmlbold),$UserName,$htmlwhite)
	$rowdata += @(,('Save as PDF',($htmlsilver -bor $htmlbold),$PDF.ToString(),$htmlwhite))
	$rowdata += @(,('Save as TEXT',($htmlsilver -bor $htmlbold),$TEXT.ToString(),$htmlwhite))
	$rowdata += @(,('Save as WORD',($htmlsilver -bor $htmlbold),$MSWORD.ToString(),$htmlwhite))
	$rowdata += @(,('Save as HTML',($htmlsilver -bor $htmlbold),$HTML.ToString(),$htmlwhite))
	$rowdata += @(,('Add DateTime',($htmlsilver -bor $htmlbold),$AddDateTime.ToString(),$htmlwhite))
	$rowdata += @(,('Hardware Inventory',($htmlsilver -bor $htmlbold),$Hardware.ToString(),$htmlwhite))
	$rowdata += @(,('Computer Name',($htmlsilver -bor $htmlbold),$ComputerName,$htmlwhite))
	$rowdata += @(,('Filename1',($htmlsilver -bor $htmlbold),$Script:FileName1,$htmlwhite))
	$rowdata += @(,('OS Detected',($htmlsilver -bor $htmlbold),$Script:RunningOS,$htmlwhite))
	$rowdata += @(,('PSUICulture',($htmlsilver -bor $htmlbold),$PSCulture,$htmlwhite))
	$rowdata += @(,('PoSH version',($htmlsilver -bor $htmlbold),$Host.Version.ToString(),$htmlwhite))
	FormatHTMLTable "Example of Horizontal AutoFitContents HTML Table" -rowArray $rowdata

	The 'rowArray' paramater is mandatory to build the table, but it is not set as such in the function - if nothing is passed, the table will be empty.

	Colors and Bold/Italics Flags are shown below:

		htmlbold       
		htmlitalics    
		htmlred        
		htmlcyan        
		htmlblue       
		htmldarkblue   
		htmllightblue   
		htmlpurple      
		htmlyellow      
		htmllime       
		htmlmagenta     
		htmlwhite       
		htmlsilver      
		htmlgray       
		htmlolive       
		htmlorange      
		htmlmaroon      
		htmlgreen       
		htmlblack     

#>

Function FormatHTMLTable
{
	Param([string]$tableheader,
	[string]$tablewidth="auto",
	[string]$fontName="Calibri",
	[int]$fontSize=2,
	[switch]$noBorder=$false,
	[int]$noHeadCols=1,
	[object[]]$rowArray=@(),
	[object[]]$fixedWidth=@(),
	[object[]]$columnArray=@())

	$HTMLBody = "<b><font face='" + $fontname + "' size='" + ($fontsize + 1) + "'>" + $tableheader + "</font></b>"

	If($columnArray.Length -eq 0)
	{
		$NumCols = $noHeadCols + 1
	}  # means we have no column headers, just a table
	Else
	{
		$NumCols = $columnArray.Length
	}  # need to add one for the color attrib

	If($Null -ne $rowArray)
	{
		$NumRows = $rowArray.length + 1
	}
	Else
	{
		$NumRows = 1
	}

	If($noBorder)
	{
		$htmlbody += "<table border='0' width='" + $tablewidth + "'>"
	}
	Else
	{
		$htmlbody += "<table border='1' width='" + $tablewidth + "'>"
	}

	If(!($columnArray.Length -eq 0))
	{
		$htmlbody += "<tr>"

		For($columnIndex = 0; $columnIndex -lt $NumCols; $columnindex+=2)
		{
			$tmp = CheckHTMLColor $columnArray[$columnIndex+1]
			If($fixedWidth.Length -eq 0)
			{
				$htmlbody += "<td style=""background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}
			Else
			{
				$htmlbody += "<td style=""width:$($fixedWidth[$columnIndex/2]); background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}

			If($columnArray[$columnIndex+1] -band $htmlbold)
			{
				$htmlbody += "<b>"
			}
			If($columnArray[$columnIndex+1] -band $htmlitalics)
			{
				$htmlbody += "<i>"
			}
			If($Null -ne $columnArray[$columnIndex])
			{
				If($columnArray[$columnIndex] -eq " " -or $columnArray[$columnIndex].length -eq 0)
				{
					$htmlbody += "&nbsp;&nbsp;&nbsp;"
				}
				Else
				{
					$found = $false
					For($i=0;$i -lt $columnArray[$columnIndex].length;$i+=2)
					{
						If($columnArray[$columnIndex][$i] -eq " ")
						{
							$htmlbody += "&nbsp;"
						}
						If($columnArray[$columnIndex][$i] -ne " ")
						{
							Break
						}
					}
					$htmlbody += $columnArray[$columnIndex]
				}
			}
			Else
			{
				$htmlbody += "&nbsp;&nbsp;&nbsp;"
			}
			If($columnArray[$columnIndex+1] -band $htmlbold)
			{
				$htmlbody += "</b>"
			}
			If($columnArray[$columnIndex+1] -band $htmlitalics)
			{
				$htmlbody += "</i>"
			}
			$htmlbody += "</font></td>"
		}
		$htmlbody += "</tr>"
	}
	$rowindex = 2
	If($Null -ne $rowArray)
	{
		AddHTMLTable $fontName $fontSize -colCount $numCols -rowCount $NumRows -rowInfo $rowArray -fixedInfo $fixedWidth
		$rowArray = @()
		$htmlbody = "</table>"
	}
	Else
	{
		$HTMLBody += "</table>"
	}	
	out-file -FilePath $Script:FileName1 -Append -InputObject $HTMLBody 4>$Null 
}
#endregion

#region other HTML functions
#***********************************************************************************************************
# CheckHTMLColor - Called from AddHTMLTable WriteHTMLLine and FormatHTMLTable
#***********************************************************************************************************
Function CheckHTMLColor
{
	Param($hash)

	If($hash -band $htmlwhite)
	{
		Return $htmlwhitemask
	}
	If($hash -band $htmlred)
	{
		Return $htmlredmask
	}
	If($hash -band $htmlcyan)
	{
		Return $htmlcyanmask
	}
	If($hash -band $htmlblue)
	{
		Return $htmlbluemask
	}
	If($hash -band $htmldarkblue)
	{
		Return $htmldarkbluemask
	}
	If($hash -band $htmllightblue)
	{
		Return $htmllightbluemask
	}
	If($hash -band $htmlpurple)
	{
		Return $htmlpurplemask
	}
	If($hash -band $htmlyellow)
	{
		Return $htmlyellowmask
	}
	If($hash -band $htmllime)
	{
		Return $htmllimemask
	}
	If($hash -band $htmlmagenta)
	{
		Return $htmlmagentamask
	}
	If($hash -band $htmlsilver)
	{
		Return $htmlsilvermask
	}
	If($hash -band $htmlgray)
	{
		Return $htmlgraymask
	}
	If($hash -band $htmlblack)
	{
		Return $htmlblackmask
	}
	If($hash -band $htmlorange)
	{
		Return $htmlorangemask
	}
	If($hash -band $htmlmaroon)
	{
		Return $htmlmaroonmask
	}
	If($hash -band $htmlgreen)
	{
		Return $htmlgreenmask
	}
	If($hash -band $htmlolive)
	{
		Return $htmlolivemask
	}
}

Function SetupHTML
{
	Write-Verbose "$(Get-Date): Setting up HTML"
    If($AddDateTime)
    {
		$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).html"
    }

    $htmlhead = "<html><head><meta http-equiv='Content-Language' content='da'><title>" + $Script:Title + "</title></head><body>"
    #echo $htmlhead > $FileName1
	out-file -FilePath $Script:FileName1 -Force -InputObject $HTMLHead 4>$Null
}
#endregion

#region Iain's Word table functions

<#
.Synopsis
	Add a table to a Microsoft Word document
.DESCRIPTION
	This function adds a table to a Microsoft Word document from either an array of
	Hashtables or an array of PSCustomObjects.

	Using this function is quicker than setting each table cell individually but can
	only utilise the built-in MS Word table autoformats. Individual tables cells can
	be altered after the table has been appended to the document (a table reference
	is returned).
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. Column headers will display the key names as defined.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -List

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. No column headers will be added, in a ListView format.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray

	This example adds table to the MS Word document, utilising all note property names
	the array of PSCustomObjects. Column headers will display the note property names.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -Columns FirstName,LastName,EmailAddress

	This example adds a table to the MS Word document, but only using the specified
	key names: FirstName, LastName and EmailAddress. If other keys are present in the
	array of Hashtables they will be ignored.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray -Columns FirstName,LastName,EmailAddress -Headers "First Name","Last Name","Email Address"

	This example adds a table to the MS Word document, but only using the specified
	PSCustomObject note properties: FirstName, LastName and EmailAddress. If other note
	properties are present in the array of PSCustomObjects they will be ignored. The
	display names for each specified column header has been overridden to display a
	custom header. Note: the order of the header names must match the specified columns.
#>

Function AddWordTable
{
	[CmdletBinding()]
	Param
	(
		# Array of Hashtable (including table headers)
		[Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$True, ParameterSetName='Hashtable', Position=0)]
		[ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Hashtable,
		# Array of PSCustomObjects
		[Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$True, ParameterSetName='CustomObject', Position=0)]
		[ValidateNotNullOrEmpty()] [PSCustomObject[]] $CustomObject,
		# Array of Hashtable key names or PSCustomObject property names to include, in display order.
		# If not supplied then all Hashtable keys or all PSCustomObject properties will be displayed.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [string[]] $Columns = $Null,
		# Array of custom table header strings in display order.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [string[]] $Headers = $Null,
		# AutoFit table behavior.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [int] $AutoFit = -1,
		# List view (no headers)
		[Switch] $List,
		# Grid lines
		[Switch] $NoGridLines,
		[Switch] $NoInternalGridLines,
		# Built-in Word table formatting style constant
		# Would recommend only $wdTableFormatContempory for normal usage (possibly $wdTableFormatList5 for List view)
		[Parameter(ValueFromPipelineByPropertyName=$True)] [int] $Format = 0
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'" -f $PSCmdlet.ParameterSetName);
		## Check if -Columns wasn't specified but -Headers were (saves some additional parameter sets!)
		If(($Null -eq $Columns) -and ($Null -ne $Headers)) 
		{
			Write-Warning "No columns specified and therefore, specified headers will be ignored.";
			$Columns = $Null;
		}
		ElseIf(($Null -ne $Columns) -and ($Null -ne $Headers)) 
		{
			## Check if number of specified -Columns matches number of specified -Headers
			If($Columns.Length -ne $Headers.Length) 
			{
				Write-Error "The specified number of columns does not match the specified number of headers.";
			}
		} ## end elseif
	} ## end Begin

	Process
	{
		## Build the Word table data string to be converted to a range and then a table later.
		[System.Text.StringBuilder] $WordRangeString = New-Object System.Text.StringBuilder;

		Switch ($PSCmdlet.ParameterSetName) 
		{
			'CustomObject' 
			{
				If($Null -eq $Columns) 
				{
					## Build the available columns from all availble PSCustomObject note properties
					[string[]] $Columns = @();
					## Add each NoteProperty name to the array
					ForEach($Property in ($CustomObject | Get-Member -MemberType NoteProperty)) 
					{ 
						$Columns += $Property.Name; 
					}
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Null -ne $Headers) 
					{
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{ 
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}

				## Iterate through each PSCustomObject
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Object in $CustomObject) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Object.$Column; 
					}
					## Use the ordered list to add each column in specified order
					[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end foreach
				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f ($CustomObject.Count));
			} ## end CustomObject

			Default 
			{   ## Hashtable
				If($Null -eq $Columns) 
				{
					## Build the available columns from all available hashtable keys. Hopefully
					## all Hashtables have the same keys (they should for a table).
					$Columns = $Hashtable[0].Keys;
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Null -ne $Headers) 
					{ 
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}
                
				## Iterate through each Hashtable
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Hash in $Hashtable) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Hash.$Column; 
					}
					## Use the ordered list to add each column in specified order
					[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end foreach

				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f $Hashtable.Count);
			} ## end default
		} ## end switch

		## Create a MS Word range and set its text to our tab-delimited, concatenated string
		Write-Debug ("$(Get-Date): `t`tBuilding table range");
		$WordRange = $Script:Doc.Application.Selection.Range;
		$WordRange.Text = $WordRangeString.ToString();

		## Create hash table of named arguments to pass to the ConvertToTable method
		$ConvertToTableArguments = @{ Separator = [Microsoft.Office.Interop.Word.WdTableFieldSeparator]::wdSeparateByTabs; }

		## Negative built-in styles are not supported by the ConvertToTable method
		If($Format -ge 0) 
		{
			$ConvertToTableArguments.Add("Format", $Format);
			$ConvertToTableArguments.Add("ApplyBorders", $True);
			$ConvertToTableArguments.Add("ApplyShading", $True);
			$ConvertToTableArguments.Add("ApplyFont", $True);
			$ConvertToTableArguments.Add("ApplyColor", $True);
			If(!$List) 
			{ 
				$ConvertToTableArguments.Add("ApplyHeadingRows", $True); 
			}
			$ConvertToTableArguments.Add("ApplyLastRow", $True);
			$ConvertToTableArguments.Add("ApplyFirstColumn", $True);
			$ConvertToTableArguments.Add("ApplyLastColumn", $True);
		}

		## Invoke ConvertToTable method - with named arguments - to convert Word range to a table
		## See http://msdn.microsoft.com/en-us/library/office/aa171893(v=office.11).aspx
		Write-Debug ("$(Get-Date): `t`tConverting range to table");
		## Store the table reference just in case we need to set alternate row coloring
		$WordTable = $WordRange.GetType().InvokeMember(
			"ConvertToTable",                               # Method name
			[System.Reflection.BindingFlags]::InvokeMethod, # Flags
			$Null,                                          # Binder
			$WordRange,                                     # Target (self!)
			([Object[]]($ConvertToTableArguments.Values)),  ## Named argument values
			$Null,                                          # Modifiers
			$Null,                                          # Culture
			([String[]]($ConvertToTableArguments.Keys))     ## Named argument names
		);

		## Implement grid lines (will wipe out any existing formatting
		If($Format -lt 0) 
		{
			Write-Debug ("$(Get-Date): `t`tSetting table format");
			$WordTable.Style = $Format;
		}

		## Set the table autofit behavior
		If($AutoFit -ne -1) 
		{ 
			$WordTable.AutoFitBehavior($AutoFit); 
		}

		If(!$List)
		{
			#the next line causes the heading row to flow across page breaks
			$WordTable.Rows.First.Headingformat = $wdHeadingFormatTrue;
		}

		If(!$NoGridLines) 
		{
			$WordTable.Borders.InbandedStyle = $wdLineStyleSingle;
			$WordTable.Borders.OutbandedStyle = $wdLineStyleSingle;
		}
		If($NoGridLines) 
		{
			$WordTable.Borders.InbandedStyle = $wdLineStyleNone;
			$WordTable.Borders.OutbandedStyle = $wdLineStyleNone;
		}
		If($NoInternalGridLines) 
		{
			$WordTable.Borders.InbandedStyle = $wdLineStyleNone;
			$WordTable.Borders.OutbandedStyle = $wdLineStyleSingle;
		}

		Return $WordTable;

	} ## end Process
}

<#
.Synopsis
	Sets the format of one or more Word table cells
.DESCRIPTION
	This function sets the format of one or more table cells, either from a collection
	of Word COM object cell references, an individual Word COM object cell reference or
	a hashtable containing Row and Column information.

	The font name, font size, bold, italic , underline and shading values can be used.
.EXAMPLE
	SetWordCellFormat -Hashtable $Coordinates -Table $TableReference -Bold

	This example sets all text to bold that is contained within the $TableReference
	Word table, using an array of hashtables. Each hashtable contain a pair of co-
	ordinates that is used to select the required cells. Note: the hashtable must
	contain the .Row and .Column key names. For example:
	@ { Row = 7; Column = 3 } to set the cell at row 7 and column 3 to bold.
.EXAMPLE
	$RowCollection = $Table.Rows.First.Cells
	SetWordCellFormat -Collection $RowCollection -Bold -Size 10

	This example sets all text to size 8 and bold for all cells that are contained
	within the first row of the table.
	Note: the $Table.Rows.First.Cells returns a collection of Word COM cells objects
	that are in the first table row.
.EXAMPLE
	$ColumnCollection = $Table.Columns.Item(2).Cells
	SetWordCellFormat -Collection $ColumnCollection -BackgroundColor 255

	This example sets the background (shading) of all cells in the table's second
	column to red.
	Note: the $Table.Columns.Item(2).Cells returns a collection of Word COM cells objects
	that are in the table's second column.
.EXAMPLE
	SetWordCellFormat -Cell $Table.Cell(17,3) -Font "Tahoma" -Color 16711680

	This example sets the font to Tahoma and the text color to blue for the cell located
	in the table's 17th row and 3rd column.
	Note: the $Table.Cell(17,3) returns a single Word COM cells object.
#>

Function SetWordCellFormat 
{
	[CmdletBinding(DefaultParameterSetName='Collection')]
	Param (
		# Word COM object cell collection reference
		[Parameter(Mandatory=$true, ValueFromPipeline=$true, ParameterSetName='Collection', Position=0)] [ValidateNotNullOrEmpty()] $Collection,
		# Word COM object individual cell reference
		[Parameter(Mandatory=$true, ParameterSetName='Cell', Position=0)] [ValidateNotNullOrEmpty()] $Cell,
		# Hashtable of cell co-ordinates
		[Parameter(Mandatory=$true, ParameterSetName='Hashtable', Position=0)] [ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Coordinates,
		# Word COM object table reference
		[Parameter(Mandatory=$true, ParameterSetName='Hashtable', Position=1)] [ValidateNotNullOrEmpty()] $Table,
		# Font name
		[Parameter()] [AllowNull()] [string] $Font = $null,
		# Font color
		[Parameter()] [AllowNull()] $Color = $null,
		# Font size
		[Parameter()] [ValidateNotNullOrEmpty()] [int] $Size = 0,
		# Cell background color
		[Parameter()] [AllowNull()] $BackgroundColor = $null,
		# Force solid background color
		[Switch] $Solid,
		[Switch] $Bold,
		[Switch] $Italic,
		[Switch] $Underline
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'." -f $PSCmdlet.ParameterSetName);
	}

	Process 
	{
		Switch ($PSCmdlet.ParameterSetName) 
		{
			'Collection' {
				ForEach($Cell in $Collection) 
				{
					If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Bold) { $Cell.Range.Font.Bold = $true; }
					If($Italic) { $Cell.Range.Font.Italic = $true; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
					If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				} # end foreach
			} # end Collection
			'Cell' 
			{
				If($Bold) { $Cell.Range.Font.Bold = $true; }
				If($Italic) { $Cell.Range.Font.Italic = $true; }
				If($Underline) { $Cell.Range.Font.Underline = 1; }
				If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
				If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
				If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
				If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
				If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
			} # end Cell
			'Hashtable' 
			{
				ForEach($Coordinate in $Coordinates) 
				{
					$Cell = $Table.Cell($Coordinate.Row, $Coordinate.Column);
					If($Bold) { $Cell.Range.Font.Bold = $true; }
					If($Italic) { $Cell.Range.Font.Italic = $true; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
					If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				}
			} # end Hashtable
		} # end switch
	} # end process
}

<#
.Synopsis
	Sets alternate row colors in a Word table
.DESCRIPTION
	This function sets the format of alternate rows within a Word table using the
	specified $BackgroundColor. This function is expensive (in performance terms) as
	it recursively sets the format on alternate rows. It would be better to pick one
	of the predefined table formats (if one exists)? Obviously the more rows, the
	longer it takes :'(

	Note: this function is called by the AddWordTable function if an alternate row
	format is specified.
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 255

	This example sets every-other table (starting with the first) row and sets the
	background color to red (wdColorRed).
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 39423 -Seed Second

	This example sets every other table (starting with the second) row and sets the
	background color to light orange (weColorLightOrange).
#>

Function SetWordTableAlternateRowColor 
{
	[CmdletBinding()]
	Param (
		# Word COM object table reference
		[Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0)] [ValidateNotNullOrEmpty()] $Table,
		# Alternate row background color
		[Parameter(Mandatory=$true, Position=1)] [ValidateNotNull()] [int] $BackgroundColor,
		# Alternate row starting seed
		[Parameter(ValueFromPipelineByPropertyName=$true, Position=2)] [ValidateSet('First','Second')] [string] $Seed = 'First'
	)

	Process 
	{
		$StartDateTime = Get-Date;
		Write-Debug ("{0}: `t`tSetting alternate table row colors.." -f $StartDateTime);

		## Determine the row seed (only really need to check for 'Second' and default to 'First' otherwise
		If($Seed.ToLower() -eq 'second') 
		{ 
			$StartRowIndex = 2; 
		}
		Else 
		{ 
			$StartRowIndex = 1; 
		}

		For($AlternateRowIndex = $StartRowIndex; $AlternateRowIndex -lt $Table.Rows.Count; $AlternateRowIndex += 2) 
		{ 
			$Table.Rows.Item($AlternateRowIndex).Shading.BackgroundPatternColor = $BackgroundColor;
		}

		## I've put verbose calls in here we can see how expensive this functionality actually is.
		$EndDateTime = Get-Date;
		$ExecutionTime = New-TimeSpan -Start $StartDateTime -End $EndDateTime;
		Write-Debug ("{0}: `t`tDone setting alternate row style color in '{1}' seconds" -f $EndDateTime, $ExecutionTime.TotalSeconds);
	}
}
#endregion

#region general script functions
Function ShowScriptOptions
{
	#updated 8-Jun-2017 with additional cover page fields
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): AddDateTime     : $($AddDateTime)"
	Write-Verbose "$(Get-Date): Company Name    : $($Script:CoName)"
	Write-Verbose "$(Get-Date): Company Address : $($CompanyAddress)"
	Write-Verbose "$(Get-Date): Company Email   : $($CompanyEmail)"
	Write-Verbose "$(Get-Date): Company Fax     : $($CompanyFax)"
	Write-Verbose "$(Get-Date): Company Phone   : $($CompanyPhone)"
	Write-Verbose "$(Get-Date): Cover Page      : $($CoverPage)"
	Write-Verbose "$(Get-Date): Dev             : $($Dev)"
	If($Dev)
	{
		Write-Verbose "$(Get-Date): DevErrorFile    : $($Script:DevErrorFile)"
	}
	Write-Verbose "$(Get-Date): Filename1       : $($Script:filename1)"
	If($PDF)
	{
		Write-Verbose "$(Get-Date): Filename2       : $($Script:filename2)"
	}
	Write-Verbose "$(Get-Date): Folder          : $($Folder)"
	Write-Verbose "$(Get-Date): From            : $($From)"
	Write-Verbose "$(Get-Date): Save As HTML    : $($HTML)"
	Write-Verbose "$(Get-Date): Save As PDF     : $($PDF)"
	Write-Verbose "$(Get-Date): Save As TEXT    : $($TEXT)"
	Write-Verbose "$(Get-Date): Save As WORD    : $($MSWORD)"
	Write-Verbose "$(Get-Date): ScriptInfo      : $($ScriptInfo)"
	Write-Verbose "$(Get-Date): Smtp Port       : $($SmtpPort)"
	Write-Verbose "$(Get-Date): Smtp Server     : $($SmtpServer)"
	Write-Verbose "$(Get-Date): Title           : $($Script:Title)"
	Write-Verbose "$(Get-Date): To              : $($To)"
	Write-Verbose "$(Get-Date): Use SSL         : $($UseSSL)"
	Write-Verbose "$(Get-Date): User Name       : $($UserName)"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): OS Detected     : $($Script:RunningOS)"
	Write-Verbose "$(Get-Date): PoSH version    : $($Host.Version)"
	Write-Verbose "$(Get-Date): PSCulture       : $($PSCulture)"
	Write-Verbose "$(Get-Date): PSUICulture     : $($PSUICulture)"
	Write-Verbose "$(Get-Date): Word language   : $($Script:WordLanguageValue)"
	Write-Verbose "$(Get-Date): Word version    : $($Script:WordProduct)"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): Script start    : $($Script:StartTime)"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
}

Function validStateProp( [object] $object, [string] $topLevel, [string] $secondLevel )
{
	#function created 8-jan-2014 by Michael B. Smith
	if( $object )
	{
		If( ( gm -Name $topLevel -InputObject $object ) )
		{
			If( ( gm -Name $secondLevel -InputObject $object.$topLevel ) )
			{
				Return $True
			}
		}
	}
	Return $False
}

Function SaveandCloseDocumentandShutdownWord
{
	#bug fix 1-Apr-2014
	#reset Grammar and Spelling options back to their original settings
	$Script:Word.Options.CheckGrammarAsYouType = $Script:CurrentGrammarOption
	$Script:Word.Options.CheckSpellingAsYouType = $Script:CurrentSpellingOption

	Write-Verbose "$(Get-Date): Save and Close document and Shutdown Word"
	If($Script:WordVersion -eq $wdWord2010)
	{
		#the $saveFormat below passes StrictMode 2
		#I found this at the following two links
		#http://blogs.technet.com/b/bshukla/archive/2011/09/27/3347395.aspx
		#http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat(v=office.14).aspx
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date): Saving DOCX file"
		}
		If($AddDateTime)
		{
			$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
			If($PDF)
			{
				$Script:FileName2 += "_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
			}
		}
		Write-Verbose "$(Get-Date): Running $($Script:WordProduct) and detected operating system $($Script:RunningOS)"
		$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
		$Script:Doc.SaveAs([REF]$Script:FileName1, [ref]$SaveFormat)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
			$Script:Doc.SaveAs([REF]$Script:FileName2, [ref]$saveFormat)
		}
	}
	ElseIf($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016)
	{
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date): Saving DOCX file"
		}
		If($AddDateTime)
		{
			$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
			If($PDF)
			{
				$Script:FileName2 += "_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
			}
		}
		Write-Verbose "$(Get-Date): Running $($Script:WordProduct) and detected operating system $($Script:RunningOS)"
		$Script:Doc.SaveAs2([REF]$Script:FileName1, [ref]$wdFormatDocumentDefault)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$Script:Doc.SaveAs([REF]$Script:FileName2, [ref]$wdFormatPDF)
		}
	}

	Write-Verbose "$(Get-Date): Closing Word"
	$Script:Doc.Close()
	$Script:Word.Quit()
	If($PDF)
	{
		[int]$cnt = 0
		While(Test-Path $Script:FileName1)
		{
			$cnt++
			If($cnt -gt 1)
			{
				Write-Verbose "$(Get-Date): Waiting another 10 seconds to allow Word to fully close (try # $($cnt))"
				Start-Sleep -Seconds 10
				$Script:Word.Quit()
				If($cnt -gt 2)
				{
					#kill the winword process

					#find out our session (usually "1" except on TS/RDC or Citrix)
					$SessionID = (Get-Process -PID $PID).SessionId
					
					#Find out if winword is running in our session
					$wordprocess = ((Get-Process 'WinWord' -ea 0)|?{$_.SessionId -eq $SessionID}).Id
					If($wordprocess -gt 0)
					{
						Write-Verbose "$(Get-Date): Attempting to stop WinWord process # $($wordprocess)"
						Stop-Process $wordprocess -EA 0
					}
				}
			}
			Write-Verbose "$(Get-Date): Attempting to delete $($Script:FileName1) since only $($Script:FileName2) is needed (try # $($cnt))"
			Remove-Item $Script:FileName1 -EA 0 4>$Null
		}
	}
	Write-Verbose "$(Get-Date): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
	If(Test-Path variable:global:word)
	{
		Remove-Variable -Name word -Scope Global 4>$Null
	}
	$SaveFormat = $Null
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	
	#is the winword process still running? kill it

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId

	#Find out if winword is running in our session
	$wordprocess = ((Get-Process 'WinWord' -ea 0)|?{$_.SessionId -eq $SessionID}).Id
	If($wordprocess -gt 0)
	{
		Write-Verbose "$(Get-Date): WinWord process is still running. Attempting to stop WinWord process # $($wordprocess)"
		Stop-Process $wordprocess -EA 0
	}
}

Function SaveandCloseTextDocument
{
	If($AddDateTime)
	{
		$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
	}

	Write-Output $Global:Output | Out-File $Script:Filename1 4>$Null
}

Function SaveandCloseHTMLDocument
{
	If($AddDateTime)
	{
		$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).html"
	}
	Out-File -FilePath $Script:FileName1 -Append -InputObject "<p></p></body></html>" 4>$Null
}

Function SetFileName1andFileName2
{
	Param([string]$OutputFileName)
	
	If($Folder -eq "")
	{
		$pwdpath = $pwd.Path
	}
	Else
	{
		$pwdpath = $Folder
	}

	If($pwdpath.EndsWith("\"))
	{
		#remove the trailing \
		$pwdpath = $pwdpath.SubString(0, ($pwdpath.Length - 1))
	}

	#set $filename1 and $filename2 with no file extension
	If($AddDateTime)
	{
		[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName)"
		If($PDF)
		{
			[string]$Script:FileName2 = "$($pwdpath)\$($OutputFileName)"
		}
	}

	If($MSWord -or $PDF)
	{
		CheckWordPreReq

		If(!$AddDateTime)
		{
			[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName).docx"
			If($PDF)
			{
				[string]$Script:FileName2 = "$($pwdpath)\$($OutputFileName).pdf"
			}
		}

		SetupWord
	}
	ElseIf($Text)
	{
		If(!$AddDateTime)
		{
			[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName).txt"
		}
		ShowScriptOptions
	}
	ElseIf($HTML)
	{
		If(!$AddDateTime)
		{
			[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName).html"
		}
		SetupHTML
		ShowScriptOptions
	}
}

Function ProcessDocumentOutput
{
	If($MSWORD -or $PDF)
	{
		SaveandCloseDocumentandShutdownWord
	}
	ElseIf($Text)
	{
		SaveandCloseTextDocument
	}
	ElseIf($HTML)
	{
		SaveandCloseHTMLDocument
	}

	$GotFile = $False

	If($PDF)
	{
		If(Test-Path "$($Script:FileName2)")
		{
			Write-Verbose "$(Get-Date): $($Script:FileName2) is ready for use"
			$GotFile = $True
		}
		Else
		{
			Write-Warning "$(Get-Date): Unable to save the output file, $($Script:FileName2)"
			Write-Error "Unable to save the output file, $($Script:FileName2)"
		}
	}
	Else
	{
		If(Test-Path "$($Script:FileName1)")
		{
			Write-Verbose "$(Get-Date): $($Script:FileName1) is ready for use"
			$GotFile = $True
		}
		Else
		{
			Write-Warning "$(Get-Date): Unable to save the output file, $($Script:FileName1)"
			Write-Error "Unable to save the output file, $($Script:FileName1)"
		}
	}
	
	#email output file if requested
	If($GotFile -and ![System.String]::IsNullOrEmpty( $SmtpServer ))
	{
		If($PDF)
		{
			$emailAttachment = $Script:FileName2
		}
		Else
		{
			$emailAttachment = $Script:FileName1
		}
		SendEmail $emailAttachment
	}
}

Function AbortScript
{
	If($MSWord -or $PDF)
	{
		$Script:Word.quit()
		Write-Verbose "$(Get-Date): System Cleanup"
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
		If(Test-Path variable:global:word)
		{
			Remove-Variable -Name word -Scope Global
		}
	}
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	Write-Verbose "$(Get-Date): Script has been aborted"
	$ErrorActionPreference = $SaveEAPreference
	Exit
}
#endregion

#region XenMobile chaptercounters
$Chapters = 14
$Chapter = 0
#endregion XenMobile chaptercounters


#region deletable functions
## If needed, you can delete the following functions ##
Function ProcessSampleText
{
	OutputSampleText
}

Function OutputSampleText
{
	If($MSWORD -or $PDF)
	{
		#we have a cover page and the page for the table of contents
		#insert a new page to start the report

		$Script:Selection.InsertNewPage()
		WriteWordLine 1 0 "XenMobile Configuration Report"
		WriteWordLine 0 0 ""
		WriteWordLine 0 0 ""
		WriteWordLine 0 0 "This document has been created using a powershell script."
		WriteWordLine 0 0 ""
		WriteWordLine 0 0 "The powershell script connect to XenMobile Server and invoke Rest API to retrieve the configuration for the following items:"
		WriteWordLine 0 0 ""
		WriteWordLine 0 1 "Certificates" "" $null 0 $true $false
		WriteWordLine 0 1 "License" "" $null 0 $true $false
		WriteWordLine 0 1 "Applications" "" $null 0 $true $false
		WriteWordLine 0 1 "MDX Applications Settings" "" $null 0 $true $false
		WriteWordLine 0 1 "NetScaler" "" $null 0 $true $false
		WriteWordLine 0 1 "LDAP" "" $null 0 $true $false
		WriteWordLine 0 1 "Devices" "" $null 0 $true $false
		WriteWordLine 0 1 "Notification Server" "" $null 0 $true $false
		WriteWordLine 0 1 "Users Groups" "" $null 0 $true $false
		WriteWordLine 0 1 "Delivery Groups" "" $null 0 $true $false
		WriteWordLine 0 1 "Enrollment Modes" "" $null 0 $true $false
		WriteWordLine 0 1 "Role-Based Access Control" "" $null 0 $true $false
		WriteWordLine 0 1 "Client Properties" "" $null 0 $true $false
		WriteWordLine 0 1 "Server Properties" "" $null 0 $true $false
		WriteWordLine 0 0 ""
		WriteWordLine 0 0 ""
		WriteWordLine 0 0 "The current version of the script is 1.02, updates are planned to be implemented in the next few weeks."
		WriteWordLine 0 0 ""
		WriteWordLine 0 0 "If you have any question, feel free to send me email: arnaud.pain@arnaud.biz"
		WriteWordLine 0 0 ""
		WriteWordLine 0 0 "You can also follow me on my blog: http://arnaudpain.com"
		WriteWordLine 0 0 ""
		WriteWordLine 0 0 "Or on Twitter: @arnaud_pain"
		WriteWordLine 0 0 ""
		WriteWordLine 0 0 "Thanks"
	}
	ElseIf($Text)
	{
		Line 0 "XenMobile Configuration Report:" $Script:Title
		Line 0 ""
		Line 0 ""
		Line 0 "This document has been created using a powershell script."
		Line 0 ""
		Line 0 "The powershell script connect to XenMobile Server and invoke Rest API to retrieve the configuration for the following items:"
		Line 0 ""
		Line 1 "Certificates" ""
		Line 1 "License" ""
		Line 1 "Applications" ""
		Line 1 "MDX Applications Settings" ""
		Line 1 "NetScaler" ""
		Line 1 "LDAP" ""
		Line 1 "Devices" ""
		Line 1 "Notification Server" ""
		Line 1 "Users Groups" ""
		Line 1 "Delivery Groups" ""
		Line 1 "Enrollment Modes" ""
		Line 1 "Role-Based Access Control" ""
		Line 1 "Client Properties" ""
		Line 1 "Server Properties" ""
		Line 0 ""
		Line 0 ""
		Line 0 "The current version of the script is 1.02, updates are planned to be implemented in the next few weeks."
		Line 0 ""
		Line 0 "If you have any question, feel free to send me email: arnaud.pain@arnaud.biz"
		Line 0 ""
		Line 0 "You can also follow me on my blog: http://arnaudpain.com"
		Line 0 ""
		Line 0 "Or on Twitter: @arnaud_pain"
		Line 0 ""
		Line 0 "Thanks"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "XenMobile Configuration Report"
		WriteHTMLLine 0 0 ""
		WriteHTMLLine 0 0 ""
		WriteHTMLLine 0 0 "This document has been created using a powershell script."
		WriteHTMLLine 0 0 ""
		WriteHTMLLine 0 0 "The powershell script connect to XenMobile Server and invoke Rest API to retrieve the configuration for the following items:"
		WriteHTMLLine 0 0 ""
		WriteHTMLLine 0 1 "Certificates" ""
		WriteHTMLLine 0 1 "License" ""
		WriteHTMLLine 0 1 "Applications" ""
		WriteHTMLLine 0 1 "MDX Applications Settings" ""
		WriteHTMLLine 0 1 "NetScaler" ""
		WriteHTMLLine 0 1 "LDAP" ""
		WriteHTMLLine 0 1 "Devices" ""
		WriteHTMLLine 0 1 "Notification Server" ""
		WriteHTMLLine 0 1 "Users Groups" ""
		WriteHTMLLine 0 1 "Delivery Groups" ""
		WriteHTMLLine 0 1 "Enrollment Modes" ""
		WriteHTMLLine 0 1 "Role-Based Access Control" ""
		WriteHTMLLine 0 1 "Client Properties" ""
		WriteHTMLLine 0 1 "Server Properties" ""
		WriteHTMLLine 0 0 ""
		WriteHTMLLine 0 0 ""
		WriteHTMLLine 0 0 "The current version of the script is 1.02, updates are planned to be implemented in the next few weeks."
		WriteHTMLLine 0 0 ""
		WriteHTMLLine 0 0 "If you have any question, feel free to send me email: arnaud.pain@arnaud.biz"
		WriteHTMLLine 0 0 ""
		WriteHTMLLine 0 0 "You can also follow me on my blog: http://arnaudpain.com"
		WriteHTMLLine 0 0 ""
		WriteHTMLLine 0 0 "Or on Twitter: @arnaud_pain"
		WriteHTMLLine 0 0 ""
		WriteHTMLLine 0 0 "Thanks"
	}
}

#region connect to XMS
Function ConnectXMS
{
	#Bypass certificate verification to enable access with XMS IP Address 
	[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

	#Login to XenMobile Server
	$Global:XMS = Read-Host -Prompt 'Please provide url of the XMS Server'

	#Function XMS-Test to verify Host exist if FQDN
	$DNSName = $XMS
	Function XMS-Test
	{
		trap [System.Management.Automation.MethodInvocationException]{
		write-host "Warning: Host does not exist" -ForegroundColor Red; 
		write-host "Please verify the address provided" -Foregroundcolor Yellow; $host.ui.RawUI.ForegroundColor = "white"; exit}
		([System.Net.Dns]::GetHostAddresses($XMS)>$null)
		$host.ui.RawUI.ForegroundColor = "Green"
		write-host "	Host verification successful"
		$host.ui.RawUI.ForegroundColor = "White"
	}

	#Define Function to check if port 4443 is opened 
	Function Port-Test
	{
		$test=(New-Object System.Net.Sockets.TcpClient).Connect($XMS, 4443) 
		trap [System.Management.Automation.MethodInvocationException]{
		write-host "Warning: Port 4443 is not opened" -ForegroundColor Red; 
		write-host "" -Foregroundcolor Yellow; $host.ui.RawUI.ForegroundColor = "white"; exit}
		$host.ui.RawUI.ForegroundColor = "Green"
		write-host "	Port 4443 is open"
		$host.ui.RawUI.ForegroundColor = "White"
	}

	#Check if XMS server can be resolved
	$host.ui.RawUI.ForegroundColor = "Yellow"
	write-host "Verifying Host:" $XMS
	$host.ui.RawUI.ForegroundColor = "white"
	XMS-Test

	#Check if port 4443 is opened 
	$host.ui.RawUI.ForegroundColor = "Yellow"
	write-host "Verifying if port 4443 is open for" $XMS
	$host.ui.RawUI.ForegroundColor = "white"
	Port-Test

	#Get Login credentials
	write-host "Please provide username and password"
	$Credential = get-credential $null

	#Check Credentials before continue
	$log = '{{"login":"{0}","password":"{1}"}}'
	$Global:cred = ($log -f $Credential.UserName, $Credential.GetNetworkCredential().Password)

	$headers=@{"Content-Type" = "application/json"}
	$Url = "https://${XMS}:4443/xenmobile/api/v1/authentication/login"
	$json=Invoke-RestMethod -Uri $url -Body $cred -Headers $headers -Method POST -Verbose:$false
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
}
#endregion

#region Certificates
Function ProcessCert
{
	#Bypass certificate verification to enable access with XMS IP Address 
	[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

	#Connect to XMS server 
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/authentication/login -Body $cred -Headers $headers -Method POST -Verbose:$false
	$headers.add("auth_token",$json.auth_token)

	#List installed Certificates 
	$cert=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/certificates" -Headers $headers -Method Get -Verbose:$false

	#Declare an array to collect our result objects
	$resultsarray =@()

	#$count will be the ‘loop counter’.
	$count = $cert.certificate.length
	$Chapter++
	Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters Certificates" 

	Try
	{
		$Script:Certificates = $cert
	}

	Catch
	{
		$Script:Certificates = $Null
	}

	If(!$? -or $Null -eq $Script:Certificates)
	{
		Write-Warning "No Certificates were retrieved."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "Warning: No Certificates were retrieved" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 0 "Warning: No Certificates were retrieved"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Warning: No Certificates were retrieved" "" $Null 0 $htmlbold
		}
		Return $False
	}
	ElseIf($? -and $Null -eq $Script:Certificates)
	{
		Write-Warning "Certificates retrieval was successful but no Certificates were returned."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "Certificates retrieval was successful but no Certificates were returned." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 0 "Certificates retrieval was successful but no Certificates were returned."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Certificates retrieval was successful but no Certificates were returned." "" $Null 0 $htmlbold
		}
		Return $False
	}
	Else
	{
		If($cert -is [array])
		{
			[int]$Script:NumCertificates = $count
		}
		Else
		{
			[int]$Script:NumCertificates = 1
		}
		Write-Verbose "$(Get-Date): `t`t$($count) Certificates found"
		Return $True
	}
}

Function ProcessCertificates
{
	If($MSWORD -or $PDF)
	{
		$Script:Selection.InsertNewPage()
		Write-Verbose "$(Get-Date): `t`t`tProcessing Certificates information"
		WriteWordLine 1 0 "Certificates"
	}
	ElseIf($Text)
	{
		Write-Verbose "$(Get-Date): `t`t`tProcessing Certificates information"
		Line 0 ""
		Line 0 "Certificates"
	}
	ElseIf($HTML)
	{
		Write-Verbose "$(Get-Date): `t`t`tProcessing Certificates	information"
		WriteHTMLLine 0 0 ""
		WriteHTMLLine 1 0 "Certificates"
		WriteHTMLLine 0 0 ""
	}
	
	OutputCertificates $Script:Certificates $Script:NumCertificates
	If($MSWORD -or $PDF)
	{
		OutputCertificatesNoInternalGridLines $Script:Certificates $Script:NumCertificates
	}
}

Function OutputCertificates
{
	#Bypass certificate verification to enable access with XMS IP Address
	[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

	#Connect to XMS server 
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/authentication/login -Body $cred -Headers $headers -Method POST -Verbose:$false
	$headers.add("auth_token",$json.auth_token)

	#List installed Certificates
	$cert=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/certificates" -Headers $headers -Method Get -Verbose:$false

	#Declare an array to collect our result objects
	$resultsarray =@()

	#$count will be the ‘loop counter’
	$count = $cert.certificate.length

	Param([object]$cert, [int]$count)
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 ""
		WriteWordLine 0 1 "$($count) Certificates found on the XMS Server $XMS"
		WriteWordLine 0 1 ""

		## IB - replacement Services table generation utilising AddWordTable function

		## Create an array of hashtables to store our services
		[System.Collections.Hashtable[]] $CertificatesWordTable = @();
		## Create an array of hashtables to store references of cells that we wish to highlight after the table has been added
		[System.Collections.Hashtable[]] $HighlightedCells = @();
		## Seed the $Services row index from the second row
		[int] $CurrentCertificateIndex = 0;
	}
	
	ElseIf($Text)
	{
		Line 0 ""
		Line 0 "`t$count Certificates found on the XMS Server $XMS"
		Line 0 ""
	}
	
	ElseIf($HTML)
	{
		$rowdata = @()
	}

	$CurrentCertificateIndex=0
	ForEach($certificate in $cert) 
	{
		for ($v=0;$v -lt $count; $v++)
		{
			If($MSWord -or $PDF)
			{
				## Add the required key/values to the hashtable
				$WordTableRowHash = @{ DisplayName = $certificate.certificate[$CurrentCertificateIndex].name;
									   Description = $certificate.certificate[$CurrentCertificateIndex].description; 
									   Remainingdays = $certificate.certificate[$CurrentCertificateIndex].remainingdays; 
									   validfrom = $certificate.certificate[$CurrentCertificateIndex].validfrom; 
									   validto = $certificate.certificate[$CurrentCertificateIndex].validto; 
									   type = $certificate.certificate[$CurrentCertificateIndex].type; 
									   isactive = $certificate.certificate[$CurrentCertificateIndex].isactive; 
									   privatekey = $certificate.certificate[$CurrentCertificateIndex].privatekey;
									 }

				## Add the hash to the array
				$CertificatesWordTable += $WordTableRowHash;
				## Store "to highlight" cell references
				If($certificate.certificate[$CurrentCertificateIndex].remainingdays -lt 31) 
				{
					$HighlightedCells += @{ Row = $CurrentCertificateIndex+2; Column = 3; }
				}
				$CurrentCertificateIndex++
			}
			ElseIf($Text)
			{
				Line 0 "Display Name`t: " $certificate.certificate[$CurrentCertificateIndex].name
				Line 0 "Description`t: " $certificate.certificate[$CurrentCertificateIndex].description
				Line 0 "Remaining Days`t: " $certificate.certificate[$CurrentCertificateIndex].remainingdays
				Line 0 "Valid From`t: " $certificate.certificate[$CurrentCertificateIndex].validfrom
				Line 0 "Valid To`t: " $certificate.certificate[$CurrentCertificateIndex].validto
				Line 0 "Type`t`t: " $certificate.certificate[$CurrentCertificateIndex].type
				Line 0 "Is Active`t: " $certificate.certificate[$CurrentCertificateIndex].isactive
				Line 0 "Private Key`t: " $certificate.certificate[$CurrentCertificateIndex].privatekey
				Line 0 ""
				$CurrentCertificateIndex++
			}
			ElseIf($HTML)
			{
				$HighlightedCells = $htmlwhite
				$rowdata += @(,(
				$certificate.certificate[$CurrentCertificateIndex].name,$htmlwhite,
				$certificate.certificate[$CurrentCertificateIndex].description,$htmlwhite,
				$certificate.certificate[$CurrentCertificateIndex].remainingdays,$htmlwhite,
				$certificate.certificate[$CurrentCertificateIndex].validfrom,$htmlwhite,
				$certificate.certificate[$CurrentCertificateIndex].validto,$htmlwhite,
				$certificate.certificate[$CurrentCertificateIndex].type,$htmlwhite,
				$certificate.certificate[$CurrentCertificateIndex].isactive,$htmlwhite,
				$certificate.certificate[$CurrentCertificateIndex].privatekey,$htmlwhite))
				$CurrentCertificateIndex++
			}	
		}
		$new++
	}

	If($MSWord -or $PDF)
	{
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		$Table = AddWordTable -Hashtable $CertificatesWordTable `
		-Columns DisplayName, Description, Remainingdays, validfrom, validto, type, isactive, privatekey `
		-Headers "Display Name", "Description", "Remaining days", "Valid from", "Valid to", "Type", "Is Active", "Private Key" `
		-Format -155 `

		## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
		## AP - Set the size for this table to 10
		SetWordCellFormat -Collection $Table -Size 10
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Color $wdColorWhite -Bold -BackgroundColor $wdColorRoyalBlue;
		##Set colums width
		$Table.Columns.Item(1).Width = 160;
		$Table.Columns.Item(2).Width = 150;
		$Table.Columns.Item(3).Width = 60;
		$Table.Columns.Item(4).Width = 60;
		$Table.Columns.Item(5).Width = 60;
		$Table.Columns.Item(6).Width = 50;
		$Table.Columns.Item(7).Width = 50;
		$Table.Columns.Item(8).Width = 50;
		
		## IB - Set the required highlighted cells
		SetWordCellFormat -Coordinates $HighlightedCells -Table $Table -Bold -BackgroundColor $wdColorRed -Solid;

		#indent the entire table 1 tab stop
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($Text)
	{
	}
	ElseIf($HTML)
	{
		$columnHeaders = @(
		'Display Name',($htmlsilver -bor $htmlbold),
		'Description',($htmlsilver -bor $htmlbold),
		'Remaining Days',($htmlsilver -bor $htmlbold),
		'Valid From',($htmlsilver -bor $htmlbold),
		'Valid To',($htmlsilver -bor $htmlbold),
		'Type',($htmlsilver -bor $htmlbold),
		'Is Active',($htmlsilver -bor $htmlbold),
		'Private Key',($htmlsilver -bor $htmlbold))
		$msg = "$count Certificates found"
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 ""
	}
}
#endregion

#region Licenses
Function ProcessLic
{
	#Bypass certificate verification to enable access with XMS IP Address 
	[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

	#Connect to XMS server
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/authentication/login -Body $cred -Headers $headers -Method POST -Verbose:$false
	$headers.add("auth_token",$json.auth_token)

	#List installed Licenses
	$lic=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/licenses" -Headers $headers -Method Get -Verbose:$false

	#Declare an array to collect our result objects
	$resultsarray =@()

	#$count will be the ‘loop counter’
	$count = $lic.cpLicenseServer.LicenseList.length
	$Chapter++
	Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters Licenses"

	Try
	{
		$Script:Licenses = $lic
	}

	Catch
	{
		$Script:Licenses = $Null
	}

	If(!$? -or $Null -eq $Script:Licenses)
	{
		If($XMS -like '*xm.citrix.com*')
		{
			Write-Warning "Your XMS is in the Cloud and no Licenses can be retrieved"
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 "Warning: Your XMS is in the Cloud and no Licenses can be retrieved." "" $Null 0 $False $True
			}
			ElseIf($Text)
			{
				Line 0 "Warning: Your XMS is in the Cloud and no Licenses can be retrieved."
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 0 0 "Warning: Your XMS is in the Cloud and no Licenses can be retrieved." "" $Null 0 $htmlbold
			}
			Return $False
		}
		Write-Warning "No Licenses were retrieved."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "Warning: No Licenses were retrieved" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 0 "Warning: No Licenses were retrieved"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Warning: No Licenses were retrieved" "" $Null 0 $htmlbold
		}
		Return $False
	}
	ElseIf($? -and $Null -eq $Script:Licenses)
	{
		Write-Warning "Licenses retrieval was successful but no Licenses were returned."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "Licenses retrieval was successful but no Licenses were returned." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 0 "Licenses retrieval was successful but no Licenses were returned."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Licenses retrieval was successful but no Licenses were returned." "" $Null 0 $htmlbold
		}
		Return $False
	}
	Else
	{
		If($lic -is [array])
		{
			[int]$Script:NumLicenses = $count
		}
		Else
		{
			[int]$Script:NumLicenses = 1
		}
		Write-Verbose "$(Get-Date): `t`t$($count) License configuration found"
		Return $True
	}
}

Function ProcessLicenses
{
	If($MSWORD -or $PDF)
	{
		$Script:Selection.InsertNewPage()
		Write-Verbose "$(Get-Date): `t`t`tProcessing Licenses information"
		WriteWordLine 1 0 "Licenses"
	}
	ElseIf($Text)
	{
		Write-Verbose "$(Get-Date): `t`t`tProcessing Licenses information"
		Line 0 ""
		Line 0 "Licenses"
	}
	ElseIf($HTML)
	{
		Write-Verbose "$(Get-Date): `t`t`tProcessing Licenses	information"
		WriteHTMLLine 0 0 ""
		WriteHTMLLine 1 0 "Licenses"
		WriteHTMLLine 0 0 ""
	}
	
	OutputLicenses $Script:Licenses $Script:NumLicenses
	If($MSWORD -or $PDF)
	{
		OutputLicensesNoInternalGridLines $Script:Licenses $Script:NumLicenses
	}
}

Function OutputLicenses
{
	#Bypass certificate verification to enable access with XMS IP Address
	[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

	#Connect to XMS server
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/authentication/login -Body $cred -Headers $headers -Method POST -Verbose:$false
	$headers.add("auth_token",$json.auth_token)

	#List License Configuration 
	$lic=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/licenses" -Headers $headers -Method Get -Verbose:$false

	#Declare an array to collect our result objects
	$resultsarray =@()

	#$count will be the ‘loop counter’
	$count = $lic.cpLicenseServer.LicenseList.length

	Param([object]$lic, [int]$count)
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 ""
		WriteWordLine 0 1 "$count Licenses server configuration found on the XMS Server $XMS"
		WriteWordLine 0 1 ""

		## IB - replacement Services table generation utilising AddWordTable function

		## Create an array of hashtables to store our services
		[System.Collections.Hashtable[]] $LicensesWordTable = @();
		## Create an array of hashtables to store references of cells that we wish to highlight after the table has been added
		[System.Collections.Hashtable[]] $HighlightedCells = @();
	}
	ElseIf($Text)
	{
		Line 0 ""
		Line 0 "`t$count Licenses server configuration found on the XMS Server $XMS"
		Line 0 ""
	}
	ElseIf($HTML)
	{
		$rowdata = @()
	}

	ForEach($cpLicenseServer in $lic) 
	{
		#Write-Verbose "$(Get-Date): `t`t`t Processing Licenses $($Service.DisplayName)";

		If($MSWord -or $PDF)
		{
			## Add the required key/values to the hashtable
			$WordTableRowHash = @{ serveraddress = $cpLicenseServer.cpLicenseServer.serveraddress;
								   servertype = $cpLicenseServer.cpLicenseServer.servertype;
								   licensetype = $cpLicenseServer.cpLicenseServer.licensetype;
								   remainingdays = $cpLicenseServer.cpLicenseServer.licenselist.remainingdays;
								   licensesinuse = $cpLicenseServer.cpLicenseServer.licenselist.licensesinuse;
								   licensesavailable = $cpLicenseServer.cpLicenseServer.licenselist.licensesavailable;
								 }

			## Add the hash to the array
			$LicensesWordTable += $WordTableRowHash;
			## Store "to highlight" cell references
		}
		ElseIf($Text)
		{
			Line 0 "Server Address`t`t: " $cpLicenseServer.cpLicenseServer.serveraddress
			Line 0 "Server Type`t`t: " $cpLicenseServer.cpLicenseServer.servertype
			Line 0 "License Type`t`t: " $cpLicenseServer.cpLicenseServer.licenseType
			Line 0 "Remaining Days`t`t: " $cpLicenseServer.cpLicenseServer.licenselist.remainingdays
			Line 0 "Licenses in Use`t`t: " $cpLicenseServer.cpLicenseServer.licenselist.licensesinuse
			Line 0 "Licenses Available`t: " $cpLicenseServer.cpLicenseServer.licenselist.licensesavailable
			Line 0 ""
		}
		ElseIf($HTML)
		{
			$HighlightedCells = $htmlwhite
			$rowdata += @(,(
			$cpLicenseServer.cpLicenseServer.serveraddress,$htmlwhite,
			$cpLicenseServer.cpLicenseServer.servertype,$htmlwhite,
			$cpLicenseServer.cpLicenseServer.licenseType,$htmlwhite,
			$cpLicenseServer.cpLicenseServer.licenselist.remainingdays,$htmlwhite,
			$cpLicenseServer.cpLicenseServer.licenselist.licensesinuse,$htmlwhite,
			$cpLicenseServer.cpLicenseServer.licenselist.licensesavailable,$htmlwhite))
		}	
	}

	If($MSWord -or $PDF)
	{
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		$Table = AddWordTable -Hashtable $LicensesWordTable `
		-Columns serveraddress, servertype, licensetype, remainingdays, licensesinuse, licensesavailable `
		-Headers "Server Address", "Server Type", "License Type", "Remaining Days", "Licenses Use", "Licenses Available" `
		-Format -155 `
		
		## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
		## AP - Set the size for this table to 10
		SetWordCellFormat -Collection $Table -Size 10
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Color $wdColorWhite -Bold -BackgroundColor $wdColorRoyalBlue;
		##Set colums width
		$Table.Columns.Item(1).Width = 140;
		$Table.Columns.Item(2).Width = 120;
		$Table.Columns.Item(3).Width = 120;
		$Table.Columns.Item(4).Width = 100;
		$Table.Columns.Item(5).Width = 80;
		$Table.Columns.Item(6).Width = 80;
		
		
		## IB - Set the required highlighted cells
		SetWordCellFormat -Coordinates $HighlightedCells -Table $Table -Bold -BackgroundColor $wdColorRed -Solid;

		#indent the entire table 1 tab stop
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)
		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($Text)
	{
	}
	ElseIf($HTML)
	{
		$columnHeaders = @(
		'Server Address',($htmlsilver -bor $htmlbold),
		'Server Type',($htmlsilver -bor $htmlbold),
		'License Type',($htmlsilver -bor $htmlbold),
		'Remaining Days',($htmlsilver -bor $htmlbold),
		'Licenses In Use',($htmlsilver -bor $htmlbold),
		'Licenses Available',($htmlsilver -bor $htmlbold))
		$msg = "$count License configuration found"
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 ""
	}
}
#endregion

#region Applications
Function ProcessApp
{
	#Bypass certificate verification to enable access with XMS IP Address 
	[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

	#Connect to XMS server
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/authentication/login -Body $cred -Headers $headers -Method POST -Verbose:$false
	$headers.add("auth_token",$json.auth_token)

	#List installed Licenses 
	$app=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/application/filter" -Body '{}' -Headers $headers -Method Post -Verbose:$false

	#Declare an array to collect our result objects
	$resultsarray =@()

	#$count will be the ‘loop counter’
	$count = ConvertFrom-Json -inputobject $app.applicationlistdata.totalmatchcount
	$chapter++
	Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters Applications"

	Try
	{
		$Script:Applications = $app
	}

	Catch
	{
		$Script:Applications = $Null
	}

	If(!$? -or $Null -eq $Script:Applications)
	{
		Write-Warning "No Applications were retrieved."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "Warning: No Applications were retrieved" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 0 "Warning: No Applications were retrieved"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Warning: No Applications were retrieved" "" $Null 0 $htmlbold
		}
		Return $False
	}
	ElseIf($? -and $Null -eq $Script:Applications)
	{
		Write-Warning "Applications retrieval was successful but no Applications were returned."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "Applications retrieval was successful but no Applications were returned." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 0 "Applications retrieval was successful but no Applications were returned."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Applications retrieval was successful but no Applications were returned." "" $Null 0 $htmlbold
		}
		Return $False
	}
	Else
	{
		If($app -is [array])
		{
			[int]$Script:NumApplications = $count
		}
		Else
		{
			[int]$Script:NumApplications = 1
		}
		Write-Verbose "$(Get-Date): `t`t$($count) Applications found"
		Return $True
	}
}

Function ProcessApplications
{
	If($MSWORD -or $PDF)
	{
		$Script:Selection.InsertNewPage()
		Write-Verbose "$(Get-Date): `t`t`tProcessing Applications information"
		WriteWordLine 1 0 "Applications"
	}
	ElseIf($Text)
	{
		Write-Verbose "$(Get-Date): `t`t`tProcessing Applications information"
		Line 0 ""
		Line 0 "Applications"
	}
	ElseIf($HTML)
	{
		Write-Verbose "$(Get-Date): `t`t`tProcessing Applications information"
		WriteHTMLLine 0 0 ""
		WriteHTMLLine 1 0 "Applications"
		WriteHTMLLine 0 0 ""
	}
	
	OutputApplications $Script:Applications $Script:NumApplications
	If($MSWORD -or $PDF)
	{
		OutputApplicationsNoInternalGridLines $Script:Applications $Script:NumApplications
	}
}

Function OutputApplications
{
	#Bypass certificate verification to enable access with XMS IP Address 
	[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

	#Connect to XMS server 
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/authentication/login -Body $cred -Headers $headers -Method POST -Verbose:$false
	$headers.add("auth_token",$json.auth_token)

	#List Applications Configuration 
	$Body=
	'
	{
	"start": 0,
	"limit": 1000
	}
	'
	$app=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/application/filter" -Body $Body -Headers $headers -Method Post -Verbose:$false

	#Declare an array to collect our result objects
	$resultsarray =@()

	#$count will be the ‘loop counter’
	$count = ConvertFrom-Json -inputobject $app.applicationlistdata.totalmatchcount

	Param([object]$app, [int]$count)
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 ""
		WriteWordLine 0 1 "$count Applications found on the XMS Server $XMS"
		WriteWordLine 0 1 ""

		## IB - replacement Services table generation utilising AddWordTable function

		## Create an array of hashtables to store our services
		[System.Collections.Hashtable[]] $ApplicationsWordTable = @();
		## Create an array of hashtables to store references of cells that we wish to highlight after the table has been added
		[System.Collections.Hashtable[]] $HighlightedCells = @();
		[int] $CurrentApplicationsIndex = 0;
	}
	
	ElseIf($Text)
	{
		Line 0 ""
		Line 0 "`t$count Applications configuration found on the XMS Server $XMS"
		Line 0 ""
	}
	
	ElseIf($HTML)
	{
		$rowdata = @()
	}

	$CurrentApplicationsIndex = 0
	for ($v=0;$v -lt $count; $v++)
	{
		ForEach($applicationlistdata in $app) 
		{
			#Write-Verbose "$(Get-Date): `t`t`t Processing Certificates $($Service.DisplayName)";

			If($MSWord -or $PDF)
			{
				## Add the required key/values to the hashtable
				$WordTableRowHash = @{ name = $applicationlistdata.applicationlistdata.applist[$CurrentApplicationsIndex].name;
									   description = $applicationlistdata.applicationlistdata.applist[$CurrentApplicationsIndex].description;
									   createdon = $applicationlistdata.applicationlistdata.applist[$CurrentApplicationsIndex].createdon;
									   lastupdate = $applicationlistdata.applicationlistdata.applist[$CurrentApplicationsIndex].lastupdated;
									   disabled = $applicationlistdata.applicationlistdata.applist[$CurrentApplicationsIndex].disabled;
									   apptype = $applicationlistdata.applicationlistdata.applist[$CurrentApplicationsIndex].apptype;
									 }

				## Add the hash to the array
				$ApplicationsWordTable += $WordTableRowHash;
				$CurrentApplicationsIndex++;
			}
			ElseIf($Text)
			{
				Line 0 "Name`t`t`t: " $applicationlistdata.applicationlistdata.applist[$CurrentApplicationsIndex].name
				Line 0 "Description`t`t: " $applicationlistdata.applicationlistdata.applist[$CurrentApplicationsIndex].description
				Line 0 "Created On`t`t: " $applicationlistdata.applicationlistdata.applist[$CurrentApplicationsIndex].createdon
				Line 0 "Last Update`t`t: " $applicationlistdata.applicationlistdata.applist[$CurrentApplicationsIndex].lastupdated
				Line 0 "Disabled`t`t: " $applicationlistdata.applicationlistdata.applist[$CurrentApplicationsIndex].disabled
				Line 0 "App Type`t`t: " $applicationlistdata.applicationlistdata.applist[$CurrentApplicationsIndex].apptype
				Line 0 ""
				$CurrentApplicationsIndex++;
			}
			ElseIf($HTML)
			{
				$HighlightedCells = $htmlwhite
				$rowdata += @(,(
				$applicationlistdata.applicationlistdata.applist[$CurrentApplicationsIndex].name,$htmlwhite,
				$applicationlistdata.applicationlistdata.applist[$CurrentApplicationsIndex].description,$htmlwhite,
				$applicationlistdata.applicationlistdata.applist[$CurrentApplicationsIndex].createdon,$htmlwhite,
				$applicationlistdata.applicationlistdata.applist[$CurrentApplicationsIndex].lastupdated,$htmlwhite,
				$applicationlistdata.applicationlistdata.applist[$CurrentApplicationsIndex].disabled,$htmlwhite,
				$applicationlistdata.applicationlistdata.applist[$CurrentApplicationsIndex].apptype,$htmlwhite))
				$CurrentApplicationsIndex++;
			}	
		}
	}

	If($MSWord -or $PDF)
	{
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		$Table = AddWordTable -Hashtable $ApplicationsWordTable `
		-Columns name, description, createdon, lastupdate, disabled, apptype `
		-Headers "Name", "Description", "Created On", "Last Update", "Disabled", "App Type" `
		-Format -155 `
		
		## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
		## AP - Set the size for this table to 8
		SetWordCellFormat -Collection $Table -Size 10
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Color $wdColorWhite -Bold -BackgroundColor $wdColorRoyalBlue;
		# Set Columns Width
		$Table.Columns.Item(1).Width = 170;
		$Table.Columns.Item(2).Width = 180;
		$Table.Columns.Item(3).Width = 90;
		$Table.Columns.Item(4).Width = 90;
		$Table.Columns.Item(5).Width = 50;
		$Table.Columns.Item(6).Width = 60;
		## IB - Set the required highlighted cells
		SetWordCellFormat -Coordinates $HighlightedCells -Table $Table -Bold -BackgroundColor $wdColorRed -Solid;

		#indent the entire table 1 tab stop
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)
		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($Text)
	{
	}
	ElseIf($HTML)
	{
		$columnHeaders = @(
		'Name',($htmlsilver -bor $htmlbold),
		'Description',($htmlsilver -bor $htmlbold),
		'Created On',($htmlsilver -bor $htmlbold),
		'Last Update',($htmlsilver -bor $htmlbold),
		'Disabled',($htmlsilver -bor $htmlbold),
		'App Type',($htmlsilver -bor $htmlbold))
		$msg = "$count Applications configuration found"
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 ""
	}
}
#endregion

#region MDX Applications
Function ProcessMDXApp
{
	#Bypass certificate verification to enable access with XMS IP Address 
	[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

	#Connect to XMS server
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/authentication/login -Body $cred -Headers $headers -Method POST -Verbose:$false
	$headers.add("auth_token",$json.auth_token)

	#List MDX Applications 
	$categories='['
	$categories=$categories + 'application.type.mdx'
	$categories=$categories + "]"

	$body=
	'
	{
	"start": 0,
	"limit": 1000,
	"filterIds":"' + $categories +'"
	}
	'
	$app1=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/application/filter" -Body $body -Headers $headers -Method Post -Verbose:$false

	#Declare an array to collect our result objects
	$resultsarray =@()

	#$count will be the ‘loop counter’
	$count1 = ConvertFrom-Json -inputobject $app1.applicationlistdata.totalmatchcount
 
	$chapter++
	Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters MDX Applications"

	Try
	{
		$Script:MDXApps = $app1
	}

	Catch
	{
		$Script:MDXApps = $Null
	}

	If(!$? -or $Null -eq $Script:MDXApps)
	{
		Write-Warning "No Applications were retrieved."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "Warning: No MDX Settings were retrieved" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 0 "Warning: No MDX Settings were retrieved"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Warning: No MDX Settings were retrieved" "" $Null 0 $htmlbold
		}
		Return $False
	}
	ElseIf($? -and $Null -eq $Script:MDXApps)
	{
		Write-Warning "MDX Applications retrieval was successful but no MDX Settings were returned."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "MDX Applications retrieval was successful but no MDX Settings were returned." "" $Null 0 $False $True
		} 
		ElseIf($Text)
		{
			Line 0 "MDX Applications retrieval was successful but no MDX Settings were returned."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "MDX Applications retrieval was successful but no MDX Settings were returned." "" $Null 0 $htmlbold
		}
		Return $False
	}
	Else
	{
		If($app1 -is [array])
		{
			[int]$Script:NumMDXApps = $count1
		}
		Else
		{
			[int]$Script:NumMDXApps = 1
		}
		Write-Verbose "$(Get-Date): `t`t$($count1) MDX Applications found"
		Return $True
	}
}

Function ProcessMDXApplications
{
	If($MSWORD -or $PDF)
	{
		$Script:Selection.InsertNewPage()
		Write-Verbose "$(Get-Date): `t`t`tProcessing MDX Applications information"
		WriteWordLine 1 0 "MDX Applications Settings"
	}
	ElseIf($Text)
	{
		Write-Verbose "$(Get-Date): `t`t`tProcessing MDX Applications information"
		Line 0 "MDX Applications Settings"
		Line 0 ""
	}
	ElseIf($HTML)
	{
		Write-Verbose "$(Get-Date): `t`t`tProcessing MDX Applications information"
		WriteHTMLLine 0 0 ""
		WriteHTMLLine 1 0 "MDX Applications Settings"
		WriteHTMLLine 0 0 ""
	}
	
	OutputMDXApplications $Script:MDXApps $Script:NumMDXApps
	If($MSWORD -or $PDF)
	{
		OutputApplicationsNoInternalGridLines $Script:MDXApps $Script:NumMDXApps
	}
}

Function OutputMDXApplications
{
	#Bypass certificate verification to enable access with XMS IP Address 
	[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

	#Connect to XMS server 
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/authentication/login -Body $cred -Headers $headers -Method POST -Verbose:$false
	$headers.add("auth_token",$json.auth_token)

	#List Applications Configuration 
	$categories='['
	$categories=$categories + 'application.type.mdx'
	$categories=$categories + "]"

	$body=
	'
	{
	"start": 0,
	"limit": 1000,
	"filterIds":"' + $categories +'"
	}
	'
	
	$app2 = Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/application/filter" -Body $Body -Headers $headers -Method Post -Verbose:$false

	#Declare an array to collect our result objects
	$resultsarray =@()
	
	$count2 = ConvertFrom-Json -inputobject $app2.applicationlistdata.totalmatchcount

	Param([object]$app2, [int]$count2)
	
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 ""
		WriteWordLine 0 1 "$count2 MDX Applications found on the XMS Server $XMS"
		WriteWordLine 0 1 ""
		## Create an array of hashtables to store our services
		[System.Collections.Hashtable[]] $Global:MDXWordTable = @();
		## Create an array of hashtables to store references of cells that we wish to highlight after the table has been added
		[System.Collections.Hashtable[]] $HighlightedCells = @();
		[int] $CurrentPoliciesIndex = 0;
		
		for($z=0;$z -lt $count2;$z++) 
		{
		WriteWordLine 2 0 $app2.applicationListData.appList[$z].name
		WriteWordLine 0 1 ""
		$id=$app2.applicationListData.appList[$z].id
		$Global:app=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/application/mobile/$id" -Headers $headers -Method Get -verbose:$false
		if(!$app.container.ios)
		{Fullfill-Android}
		else{Fullfill-iOS}
		$Script:Selection.InsertNewPage()
		$Table = $Null
		}
	}
	
	ElseIf($Text)
	{
		for($z=0;$z -lt $count2;$z++) 
		{
		Line 0 $app2.applicationListData.appList[$z].name
		Line 0 ""
		$id=$app2.applicationListData.appList[$z].id
		$Global:app=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/application/mobile/$id" -Headers $headers -Method Get -verbose:$false
		if(!$app.container.ios)
		{Fullfill-Android}
		else{Fullfill-iOS}
		Line 0 ""
		}
	}
	
	ElseIf($HTML)
	{
	for($z=0;$z -lt $count2;$z++) 
		{
		$rowdata = @()
		$id=$app2.applicationListData.appList[$z].id
		$Global:app=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/application/mobile/$id" -Headers $headers -Method Get -verbose:$false
		if(!$app.container.ios)
		{Fullfill-Android}
		else{Fullfill-iOS}
		}
	}
	
}

Function Fullfill-Android
{
	$count = $app.container.android.policies.length
	$CurrentPoliciesIndex = 0
			If($MSWord -or $PDF)
			{
			for ($v=0;$v -lt $count; $v++)
			{
				$WordTableRowHash = @{ Title = $app.container.android.policies[$CurrentPoliciesIndex].title;
									   Value = $app.container.android.policies[$CurrentPoliciesIndex].policyValue;
									 }
				$Global:MDXWordTable += $WordTableRowHash;
				$CurrentPoliciesIndex++;
			}
			}
			ElseIf($Text)
			{
			for ($v=0;$v -lt $count; $v++)
			{
				Line 0 $app.container.android.policies[$CurrentPoliciesIndex].title":"
				Line 0 "`t`t`t`t"$app.container.android.policies[$CurrentPoliciesIndex].policyValue
				$CurrentPoliciesIndex++;
			}
			}
			ElseIf($HTML)
			{
			for ($v=0;$v -lt $count; $v++)
			{
				$HighlightedCells = $htmlwhite
				$rowdata += @(,(
				$app.container.android.policies[$CurrentPoliciesIndex].title,$htmlwhite,
				$app.container.android.policies[$CurrentPoliciesIndex].policyValue,$htmlwhite))
				$CurrentPoliciesIndex++;
			}
			}			

	
	If($MSWord -or $PDF)
	{
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		$Table = AddWordTable -Hashtable $MDXWordTable `
		-Columns Title, Value `
		-Headers "Title", "Value" `
		-Format -155 `
		
		## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
		## AP - Set the size for this table to 8
		SetWordCellFormat -Collection $Table -Size 10
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Color $wdColorWhite -Bold -BackgroundColor $wdColorRoyalBlue;
		# Set Columns Width
		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 440;
		# IB - Set the required highlighted cells
		SetWordCellFormat -Coordinates $HighlightedCells -Table $Table -Bold -BackgroundColor $wdColorRed -Solid;

		#indent the entire table 1 tab stop
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)
		FindWordDocumentEnd
		#$Table = $Null
	}
	ElseIf($Text)
	{
	}
	ElseIf($HTML)
	{
		$columnHeaders = @(
		'Title',($htmlsilver -bor $htmlbold),
		'Value',($htmlsilver -bor $htmlbold))
		$msg = $app.container.name
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 ""
	}
}

Function Fullfill-iOS
{
	$count = $app.container.ios.policies.length
	$CurrentPoliciesIndex = 0
			If($MSWord -or $PDF)
			{
			for ($v=0;$v -lt $count; $v++)
			{
				$WordTableRowHash = @{ Title = $app.container.ios.policies[$CurrentPoliciesIndex].title;
									   Value = $app.container.ios.policies[$CurrentPoliciesIndex].policyValue;
									 }
				$Global:MDXWordTable += $WordTableRowHash;
				$CurrentPoliciesIndex++;
			}
			}
			ElseIf($Text)
			{
			for ($v=0;$v -lt $count; $v++)
			{
				Line 0 $app.container.ios.policies[$CurrentPoliciesIndex].title":"
				Line 0 "`t`t`t`t"$app.container.ios.policies[$CurrentPoliciesIndex].policyValue
				$CurrentPoliciesIndex++;
			}
			}
			ElseIf($HTML)
			{
			for ($v=0;$v -lt $count; $v++)
			{
				$HighlightedCells = $htmlwhite
				$rowdata += @(,(
				$app.container.ios.policies[$CurrentPoliciesIndex].title,$htmlwhite,
				$app.container.ios.policies[$CurrentPoliciesIndex].policyValue,$htmlwhite))
				$CurrentPoliciesIndex++;
			}
			}			

	
	If($MSWord -or $PDF)
	{
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		$Table = AddWordTable -Hashtable $MDXWordTable `
		-Columns Title, Value `
		-Headers "Title", "Value" `
		-Format -155 `
		
		## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
		## AP - Set the size for this table to 8
		SetWordCellFormat -Collection $Table -Size 10
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Color $wdColorWhite -Bold -BackgroundColor $wdColorRoyalBlue;
		# Set Columns Width
		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 440;
		# IB - Set the required highlighted cells
		SetWordCellFormat -Coordinates $HighlightedCells -Table $Table -Bold -BackgroundColor $wdColorRed -Solid;

		#indent the entire table 1 tab stop
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)
		FindWordDocumentEnd
		#$Table = $Null
	}
	ElseIf($Text)
	{
	}
	ElseIf($HTML)
	{
		$columnHeaders = @(
		'Title',($htmlsilver -bor $htmlbold),
		'Value',($htmlsilver -bor $htmlbold))
		$msg = $app.container.name
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 ""
	}
}
#endregion

#region NetScaler
Function ProcessNS
{
	#Bypass certificate verification to enable access with XMS IP Address 
	[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

	#Connect to XMS server
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/authentication/login -Body $cred -Headers $headers -Method POST -Verbose:$false
	$headers.add("auth_token",$json.auth_token)

	#List NetScaler configuration 
	$netscaler=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/netscaler" -Headers $headers -Method Get -Verbose:$false

	# Declare an array to collect our result objects
	$resultsarray =@()

	#$count will be the ‘loop counter’
	$count = $netscaler.aglist.length
	$Chapter++
	Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler"

	Try
	{
		$Script:NetScaler = $netscaler
	}

	Catch
	{
		$Script:NetScaler = $Null
	}

	If(!$? -or $Null -eq $Script:NetScaler)
	{
		Write-Warning "No NetScaler were retrieved."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "Warning: No NetScaler were retrieved" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 0 "Warning: No NetScaler were retrieved"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Warning: No NetScaler were retrieved" "" $Null 0 $htmlbold
		}
		Return $False
	}
	ElseIf($? -and $Null -eq $Script:NetScaler)
	{
		Write-Warning "NetScaler retrieval was successful but no NetScaler were returned."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "NetScaler retrieval was successful but no NetScaler were returned." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 0 "NetScaler retrieval was successful but no NetScaler were returned."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "NetScaler retrieval was successful but no NetScaler were returned." "" $Null 0 $htmlbold
		}
		Return $False
	}
	Else
	{
		If($netscaler -is [array])
		{
			[int]$Script:NumNetScaler = $count
		}
		Else
		{
			[int]$Script:NumNetScaler = 1
		}
		Write-Verbose "$(Get-Date): `t`t$($count) NetScaler configuration found"
		Return $True
	}
}

Function ProcessNetScaler
{
	If($MSWORD -or $PDF)
	{
		$Script:Selection.InsertNewPage()
		Write-Verbose "$(Get-Date): `t`t`tProcessing NetScaler information"
		WriteWordLine 1 0 "NetScaler"
	}
	ElseIf($Text)
	{
		Write-Verbose "$(Get-Date): `t`t`tProcessing NetScaler information"
		Line 0 ""
		Line 0 "NetScaler"
	}
	ElseIf($HTML)
	{
		Write-Verbose "$(Get-Date): `t`t`tProcessing NetScaler information"
		WriteHTMLLine 0 0 ""
		WriteHTMLLine 1 0 "NetScaler"
		WriteHTMLLine 0 0 ""
	}
	
	OutputNetScaler $Script:NetScaler $Script:NumNetScaler
	If($MSWORD -or $PDF)
	{
		OutputNetScalerNoInternalGridLines $Script:NetScaler $Script:NumNetScaler
	}
}

Function OutputNetScaler
{
	#Bypass certificate verification to enable access with XMS IP Address 
	[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

	#Connect to XMS server
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/authentication/login -Body $cred -Headers $headers -Method POST -Verbose:$false
	$headers.add("auth_token",$json.auth_token)

	#List NetScaler Configuration 
	$netscaler=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/netscaler" -Headers $headers -Method Get -Verbose:$false

	#Declare an array to collect our result objects
	$resultsarray =@()

	#$count will be the ‘loop counter’
	$count = $netscaler.aglist.length

	Param([object]$netscaler, [int]$count)
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 ""
		WriteWordLine 0 1 "$count NetScaler found on the XMS Server $XMS"
		WriteWordLine 0 1 ""

		## IB - replacement Services table generation utilising AddWordTable function

		## Create an array of hashtables to store our services
		[System.Collections.Hashtable[]] $NetScalerWordTable = @();
		## Create an array of hashtables to store references of cells that we wish to highlight after the table has been added
		[System.Collections.Hashtable[]] $HighlightedCells = @();
		[int] $CurrentNetScalerIndex = 0;
	}
	ElseIf($Text)
	{
		Line 0 ""
		Line 0 "`t$count NetScaler configuration found on the XMS Server $XMS"
		Line 0 ""
	}
	ElseIf($HTML)
	{
		$rowdata = @()
	}

	$CurrentNetScalerIndex = 0
	ForEach($aglist in $netscaler) 
	{
		for ($v=0;$v -lt $count; $v++)
		{
			#Write-Verbose "$(Get-Date): `t`t`t Processing Certificates $($Service.DisplayName)";

			If($MSWord -or $PDF)
			{
				## Add the required key/values to the hashtable
				$WordTableRowHash = @{ name = $aglist.aglist[$v].name;
									   alias = $aglist.aglist[$v].alias;
									   url = $aglist.aglist[$v].url;
									   passwordrequired = $aglist.aglist[$v].passwordrequired;
									   logontype = $aglist.aglist[$v].logontype;
									   #callbackurl = $aglist.aglist[$v].callback[$v].callbackUrl;
									   #callbackip = $aglist.aglist[$v].callback[$v].ip;
									   default = $aglist.aglist[$v].default;
									 }

				## Add the hash to the array
				$NetScalerWordTable += $WordTableRowHash;
				$CurrentNetScalerIndex++;
			}
			ElseIf($Text)
			{
				Line 0 "Name`t`t`t: " $aglist.aglist[$v].name
				Line 0 "Alias`t`t`t: " $aglist.aglist[$v].alias
				Line 0 "URL`t`t`t: " $aglist.aglist[$v].url
				Line 0 "Password Required`t: " $aglist.aglist[$v].passwordrequired
				Line 0 "Logon Type`t`t: " $aglist.aglist[$v].logontype
				Line 0 "Callback URL`t`t: " $aglist.aglist[$v].callback[$v].callbackUrl
				Line 0 "Callback IP`t`t: " $aglist.aglist[$v].callback[$v].ip
				Line 0 "Default`t`t`t: " $aglist.aglist[$v].default
				Line 0 ""
				$CurrentNetScalerIndex++;
			}
			ElseIf($HTML)
			{
				$HighlightedCells = $htmlwhite
				$rowdata += @(,(
				$aglist.aglist[$v].name,$htmlwhite,
				$aglist.aglist[$v].alias,$htmlwhite,
				$aglist.aglist[$v].url,$htmlwhite,
				$aglist.aglist[$v].passwordrequired,$htmlwhite,
				$aglist.aglist[$v].logontype,$htmlwhite,
				$aglist.aglist[$v].default,$htmlwhite))
				$CurrentNetScalerIndex++;
			}	
		}
	}

	If($MSWord -or $PDF)
	{
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		$Table = AddWordTable -Hashtable $NetScalerWordTable `
		-Columns name, alias, url, passwordrequired, logontype, default `
		-Headers "Name", "Alias", "URL", "Password Required", "Logon Type", "Default" `
		-Format -155 `

		## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
		SetWordCellFormat -Collection $Table -Size 10
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Color $wdColorWhite -Bold -BackgroundColor $wdColorRoyalBlue;
		# Set Columns Width
		$Table.Columns.Item(1).Width = 170;
		$Table.Columns.Item(2).Width = 60;
		$Table.Columns.Item(3).Width = 200;
		$Table.Columns.Item(4).Width = 70;
		$Table.Columns.Item(5).Width = 70;
		$Table.Columns.Item(6).Width = 70;
		## IB - Set the required highlighted cells
		SetWordCellFormat -Coordinates $HighlightedCells -Table $Table -Bold -BackgroundColor $wdColorRed -Solid;

		#indent the entire table 1 tab stop
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)
		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($Text)
	{
	}
	ElseIf($HTML)
	{
		$columnHeaders = @(
		'Name',($htmlsilver -bor $htmlbold),
		'Alias',($htmlsilver -bor $htmlbold),
		'URL',($htmlsilver -bor $htmlbold),
		'Password Required',($htmlsilver -bor $htmlbold),
		'Logon Type',($htmlsilver -bor $htmlbold),
		'Default',($htmlsilver -bor $htmlbold))
		$msg = "$count NetScaler configuration found"
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 ""
	}
}
#endregion

#region LDAP
Function ProcessAD
{
	#Bypass certificate verification to enable access with XMS IP Address 
	[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

	# Connect to XMS server
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/authentication/login -Body $cred -Headers $headers -Method POST -Verbose:$false
	$headers.add("auth_token",$json.auth_token)

	#List LDAP configuration 
	$ldap=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/ldap -Headers $headers -Method Get -Verbose:$false

	# Declare an array to collect our result objects
	$resultsarray =@()

	#$count will be the ‘loop counter’
	$count=$ldap.adlist.length
	$Chapter++
	Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters LDAP"

	Try
	{
		$Script:LDAP = $ldap
	}

	Catch
	{
		$Script:LDAP = $Null
	}

	If(!$? -or $Null -eq $Script:LDAP)
	{
		Write-Warning "No LDAP were retrieved."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "Warning: No LDAP were retrieved" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 0 "Warning: No LDAP were retrieved"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Warning: No LDAP were retrieved" "" $Null 0 $htmlbold
		}
		Return $False
	}
	ElseIf($? -and $Null -eq $Script:LDAP)
	{
		Write-Warning "LDAP retrieval was successful but no LDAP were returned."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "LDAP retrieval was successful but no LDAP were returned." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 0 "LDAP retrieval was successful but no LDAP were returned."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "LDAP retrieval was successful but no LDAP were returned." "" $Null 0 $htmlbold
		}
		Return $False
	}
	Else
	{
		If($ldap -is [array])
		{
			[int]$Script:NumLDAP = $count
		}
		Else
		{
			[int]$Script:NumLDAP = 1
		}
		Write-Verbose "$(Get-Date): `t`t$($count) LDAP configuration found"
		Return $True
	}
}

Function ProcessLDAP
{
	If($MSWORD -or $PDF)
	{
		$Script:Selection.InsertNewPage()
		Write-Verbose "$(Get-Date): `t`t`tProcessing LDAP information"
		WriteWordLine 1 0 "LDAP"
	}
	ElseIf($Text)
	{
		Write-Verbose "$(Get-Date): `t`t`tProcessing LDAP information"
		Line 0 ""
		Line 0 "LDAP"
	}
	ElseIf($HTML)
	{
		Write-Verbose "$(Get-Date): `t`t`tProcessing LDAP	information"
		WriteHTMLLine 0 0 ""
		WriteHTMLLine 1 0 "LDAP"
		WriteHTMLLine 0 0 ""
	}
	
	OutputLDAP $Script:LDAP $Script:NumLDAP
	If($MSWORD -or $PDF)
	{
		OutputLDAPNoInternalGridLines $Script:LDAP $Script:NumLDAP
	}
}

Function OutputLDAP
{
	#Bypass certificate verification to enable access with XMS IP Address 
	[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

	#Connect to XMS server
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/authentication/login -Body $cred -Headers $headers -Method POST -Verbose:$false
	$headers.add("auth_token",$json.auth_token)

	#List LDAP Configuration
	$ldap=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/ldap -Headers $headers -Method Get -Verbose:$false

	#Declare an array to collect our result objects
	$resultsarray =@()

	#$count will be the ‘loop counter’
	$count=$ldap.adlist.length

	Param([object]$ldap, [int]$count)
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 ""
		WriteWordLine 0 1 "$count LDAP found on the XMS Server $XMS"
		WriteWordLine 0 1 ""

		## IB - replacement Services table generation utilising AddWordTable function

		## Create an array of hashtables to store our services
		[System.Collections.Hashtable[]] $LDAPWordTable = @();
		## Create an array of hashtables to store references of cells that we wish to highlight after the table has been added
		[System.Collections.Hashtable[]] $HighlightedCells = @();
		[int] $CurrentLDAPIndex = 0;
	}
	ElseIf($Text)
	{
		Line 0 ""
		Line 0 "`t$count LDAP configuration found on the XMS Server $XMS"
		Line 0 ""
	}
	ElseIf($HTML)
	{
		$rowdata = @()
	}

	ForEach($adlist in $ldap) 
	{
		for ($v=0;$v -lt $count; $v++)
		{
			If($MSWord -or $PDF)
			{
				## Add the required key/values to the hashtable
				$WordTableRowHash = @{ domain = $adlist.adlist[$v].domain;
									   userbasedn = $adlist.adlist[$v].userbasedn;
									   groupbasedn = $adlist.adlist[$v].groupbasedn;
									   port = $adlist.adlist[$v].port;
									   username = $adlist.adlist[$v].username;
									   primaryhost = $adlist.adlist[$v].primaryhost;
									   secondaryhost = $adlist.adlist[$v].secondaryhost;
									   usesecure = $adlist.adlist[$v].usesecure;
									   globalcatalogport = $adlist.adlist[$v].globalcatalogport;
									   usersearchby = $adlist.adlist[$v].usersearchby;
									   domainalias = $adlist.adlist[$v].domainalias;
									 }

				## Add the hash to the array
				$LDAPWordTable += $WordTableRowHash;
				$CurrentLDAPIndex++;
			}
			ElseIf($Text)
			{
				Line 0 "Domain`t`t`t: " $adlist.adlist[$v].domain
				Line 0 "User Base DN`t`t: " $adlist.adlist[$v].userbasedn
				Line 0 "Group Base DN`t`t: " $adlist.adlist[$v].groupbasedn
				Line 0 "Port`t`t`t: " $adlist.adlist[$v].port
				Line 0 "Username`t`t: " $adlist.adlist[$v].username
				Line 0 "Primary Host`t`t: " $adlist.adlist[$v].primaryhost
				Line 0 "Secondary Host`t`t: " $adlist.adlist[$v].secondaryhost
				Line 0 "Use Secure`t`t: " $adlist.adlist[$v].usesecure
				Line 0 "Global Catalog Port`t: " $adlist.adlist[$v].globalcatalogport
				Line 0 "User Search By`t`t: " $adlist.adlist[$v].usersearchby
				Line 0 "Domain Alias`t`t: " $adlist.adlist[$v].domainalias
				Line 0 ""
				$CurrentLDAPIndex++;
			}
			ElseIf($HTML)
			{
				$HighlightedCells = $htmlwhite
				$rowdata += @(,(
				$adlist.adlist[$v].domain,$htmlwhite,
				$adlist.adlist[$v].userbasedn,$htmlwhite,
				$adlist.adlist[$v].groupbasedn,$htmlwhite,
				$adlist.adlist[$v].port,$htmlwhite,
				$adlist.adlist[$v].username,$htmlwhite,
				$adlist.adlist[$v].primaryhost,$htmlwhite,
				$adlist.adlist[$v].secondaryhost,$htmlwhite,
				$adlist.adlist[$v].usesecure,$htmlwhite,
				$adlist.adlist[$v].globalcatalogport,$htmlwhite,
				$adlist.adlist[$v].usersearchby,$htmlwhite,
				$adlist.adlist[$v].domainalias,$htmlwhite))
				$CurrentLDAPIndex++;
			}	
		}
	}

	If($MSWord -or $PDF)
	{
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		$Table = AddWordTable -Hashtable $LDAPWordTable `
		-Columns domain, userbasedn, groupbasedn, port, username, primaryhost, secondaryhost, usesecure, globalcatalogport, usersearchby, domainalias `
		-Headers "Domain", "User Base DN", "Group Base DN", "Port", "Username", "Primary Host", "Secondary Host", "Use Secure", "Global Catalog Port", "User Search By", "Domain Alias" `
		-Format -155 `
		
		## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
		## AP - Set the size for this table to 7
		SetWordCellFormat -Collection $Table -Size 9
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Color $wdColorWhite -Bold -BackgroundColor $wdColorRoyalBlue;
		# Set Columns Width
		$Table.Columns.Item(1).Width = 65;
		$Table.Columns.Item(2).Width = 65;
		$Table.Columns.Item(3).Width = 65;
		$Table.Columns.Item(4).Width = 30;
		$Table.Columns.Item(5).Width = 60;
		$Table.Columns.Item(6).Width = 75;
		$Table.Columns.Item(7).Width = 75;
		$Table.Columns.Item(8).Width = 40;
		$Table.Columns.Item(9).Width = 45;
		$Table.Columns.Item(10).Width = 50;
		$Table.Columns.Item(11).Width = 70;
		## IB - Set the required highlighted cells
		SetWordCellFormat -Coordinates $HighlightedCells -Table $Table -Bold -BackgroundColor $wdColorRed -Solid;

		#indent the entire table 1 tab stop
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)
		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($Text)
	{
	}
	ElseIf($HTML)
	{
		$columnHeaders = @(
		'Domain',($htmlsilver -bor $htmlbold),
		'User Base DN',($htmlsilver -bor $htmlbold),
		'Group Base DN',($htmlsilver -bor $htmlbold),
		'Port',($htmlsilver -bor $htmlbold),
		'Username',($htmlsilver -bor $htmlbold),
		'Primary Host',($htmlsilver -bor $htmlbold),
		'Secondary Host',($htmlsilver -bor $htmlbold),
		'Use Secure',($htmlsilver -bor $htmlbold),
		'Global Catalog Port',($htmlsilver -bor $htmlbold),
		'User Search By',($htmlsilver -bor $htmlbold),
		'Domain Alias',($htmlsilver -bor $htmlbold))
		$msg = "$count LDAP configuration found"
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 ""
	}
}
#endregion

#region Devices
Function ProcessDev
{
	#Bypass certificate verification to enable access with XMS IP Address 
	[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

	#Connect to XMS server
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/authentication/login -Body $cred -Headers $headers -Method POST -Verbose:$false
	$headers.add("auth_token",$json.auth_token)

	#List Devices configuration 
	$dev=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/device/filter -Body '{}' -Headers $headers -Method Post -Verbose:$false

	# Declare an array to collect our result objects
	$resultsarray =@()

	#$count will be the ‘loop counter’
	$count = $dev.matchedRecords

	$Chapter++
	Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters Devices"

	Try
	{
		$Script:Devices = $dev
	}

	Catch
	{
		$Script:Devices = $Null
	}

	If(!$? -or $Null -eq $Script:Devices)
	{
		Write-Warning "No Devices were retrieved."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "Warning: No Devices were retrieved" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 0 "Warning: No Devices were retrieved"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Warning: No Devices were retrieved" "" $Null 0 $htmlbold
		}
		Return $False
	}
	ElseIf($? -and $Null -eq $Script:Devices)
	{
		Write-Warning "Devices retrieval was successful but no Devices were returned."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "Devices retrieval was successful but no Devices were returned." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 0 "Devices retrieval was successful but no Devices were returned."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Devices retrieval was successful but no Devices were returned." "" $Null 0 $htmlbold
		}
		Return $False
	}
	Else
	{
		If($dev -is [array])
		{
			[int]$Script:NumDevices = $count
		}
		Else
		{
			[int]$Script:NumDevices = 1
		}
		Write-Verbose "$(Get-Date): `t`t$($count) Devices configuration found"
		Return $True
	}
}

Function ProcessDevices
{
	If($MSWORD -or $PDF)
	{
		$Script:Selection.InsertNewPage()
		Write-Verbose "$(Get-Date): `t`t`tProcessing Devices information"
		WriteWordLine 1 0 "Devices"
	}
	ElseIf($Text)
	{
		Write-Verbose "$(Get-Date): `t`t`tProcessing Devices information"
		Line 0 ""
		Line 0 "Devices"
	}
	ElseIf($HTML)
	{
		Write-Verbose "$(Get-Date): `t`t`tProcessing Devices	information"
		WriteHTMLLine 0 0 ""
		WriteHTMLLine 1 0 "Devices"
		WriteHTMLLine 0 0 ""
	}
	
	OutputDevices $Script:Devices $Script:NumDevices
	If($MSWORD -or $PDF)
	{
		OutputDevicesNoInternalGridLines $Script:Devices $Script:NumDevices
	}
}

Function OutputDevices
{
	#Bypass certificate verification to enable access with XMS IP Address
	[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

	#Connect to XMS server
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/authentication/login -Body $cred -Headers $headers -Method POST -Verbose:$false
	$headers.add("auth_token",$json.auth_token)

	#List Devices Configuration
	$devBody=
	'
	{
	"start": 0,
	"limit": 10000
	}
	'
	$dev=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/device/filter -Body $devBody -Headers $headers -Method Post -Verbose:$false

	#Declare an array to collect our result objects
	$resultsarray =@()

	#$count will be the ‘loop counter’
	$count = $dev.matchedRecords

	Param([object]$dev, [int]$count)
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 ""
		WriteWordLine 0 1 "$count Devices found on the XMS Server $XMS"
		WriteWordLine 0 1 ""

		## IB - replacement Services table generation utilising AddWordTable function

		## Create an array of hashtables to store our services
		[System.Collections.Hashtable[]] $DevicesWordTable = @();
		## Create an array of hashtables to store references of cells that we wish to highlight after the table has been added
		[System.Collections.Hashtable[]] $HighlightedCells = @();
		[int] $CurrentDevicesIndex = 0;
	}
	
	ElseIf($Text)
	{
		Line 0 ""
		Line 0 "`t$count Devices configuration found on the XMS Server $XMS"
		Line 0 ""
	}
	
	ElseIf($HTML)
	{
		$rowdata = @()
	}

	$CurrentDevicesIndex=0
	ForEach($devices in $dev.filteredDevicesDataList) 
	{
		#Write-Verbose "$(Get-Date): `t`t`t Processing Devices $($Service.DisplayName)";

		If($MSWord -or $PDF)
		{
			## Add the required key/values to the hashtable
			$WordTableRowHash = @{ id = $dev.filteredDevicesDataList[$CurrentDevicesIndex].id;
								   platform = $dev.filteredDevicesDataList[$CurrentDevicesIndex].platform;
								   devicemodel = $dev.filteredDevicesDataList[$CurrentDevicesIndex].devicemodel;
								   managed = $dev.filteredDevicesDataList[$CurrentDevicesIndex].managed;
								   mamregistered = $dev.filteredDevicesDataList[$CurrentDevicesIndex].mamregistered;
								   mdmknown = $dev.filteredDevicesDataList[$CurrentDevicesIndex].mdmknown;
								   devicename = $dev.filteredDevicesDataList[$CurrentDevicesIndex].devicename;
								   username = $dev.filteredDevicesDataList[$CurrentDevicesIndex].username;
								 }
			
			## Add the hash to the array
			$DevicesWordTable += $WordTableRowHash;
			$CurrentDevicesIndex++
		}
		ElseIf($Text)
		{
			Line 0 "Device ID`t`t: " $dev.filteredDevicesDataList[$CurrentDevicesIndex].id
			Line 0 "Device Platform`t`t: " $dev.filteredDevicesDataList[$CurrentDevicesIndex].platform
			Line 0 "Device Model`t`t: " $dev.filteredDevicesDataList[$CurrentDevicesIndex].devicemodel
			Line 0 "Device Managed`t`t: " $dev.filteredDevicesDataList[$CurrentDevicesIndex].managed
			Line 0 "Device MAM`t`t: " $dev.filteredDevicesDataList[$CurrentDevicesIndex].mamregistered
			Line 0 "Device MDM `t`t: " $dev.filteredDevicesDataList[$CurrentDevicesIndex].mdmknown
			Line 0 "Device Name`t`t: " $dev.filteredDevicesDataList[$CurrentDevicesIndex].devicename
			Line 0 "Enrolled User`t`t: " $dev.filteredDevicesDataList[$CurrentDevicesIndex].username
			Line 0 ""
			$CurrentDevicesIndex++
		}
		ElseIf($HTML)
		{
			$HighlightedCells = $htmlwhite
			$rowdata += @(,(
			$dev.filteredDevicesDataList[$CurrentDevicesIndex].id,$htmlwhite,
			$dev.filteredDevicesDataList[$CurrentDevicesIndex].platform,$htmlwhite,
			$dev.filteredDevicesDataList[$CurrentDevicesIndex].devicemodel,$htmlwhite,
			$dev.filteredDevicesDataList[$CurrentDevicesIndex].managed,$htmlwhite,
			$dev.filteredDevicesDataList[$CurrentDevicesIndex].mamregistered,$htmlwhite,
			$dev.filteredDevicesDataList[$CurrentDevicesIndex].mdmknown,$htmlwhite,
			$dev.filteredDevicesDataList[$CurrentDevicesIndex].devicename,$htmlwhite,
			$dev.filteredDevicesDataList[$CurrentDevicesIndex].username,$htmlwhite))
			$CurrentDevicesIndex++
		}
	}

	If($MSWord -or $PDF)
	{
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		$Table = AddWordTable -Hashtable $DevicesWordTable `
		-Columns id, platform, devicemodel, managed, mamregistered, mdmknown, devicename, username `
		-Headers "Device ID", "Device Platform", "Device Model", "Device Managed", "Device MAM", "Device MDM", "Device Name", "Enrolled User" `
		-Format -155 `
		
		## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
		SetWordCellFormat -Collection $Table -Size 10
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Color $wdColorWhite -Bold -BackgroundColor $wdColorRoyalBlue;
		# Set Columns Width
		$Table.Columns.Item(1).Width = 40;
		$Table.Columns.Item(2).Width = 50;
		$Table.Columns.Item(3).Width = 80;
		$Table.Columns.Item(4).Width = 50;
		$Table.Columns.Item(5).Width = 50;
		$Table.Columns.Item(6).Width = 50;
		$Table.Columns.Item(7).Width = 140;
		$Table.Columns.Item(8).Width = 180;
		## IB - Set the required highlighted cells
		SetWordCellFormat -Coordinates $HighlightedCells -Table $Table -Bold -BackgroundColor $wdColorRed -Solid;

		#indent the entire table 1 tab stop
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)
		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($Text)
	{
	}
	ElseIf($HTML)
	{
		$columnHeaders = @(
		'Device ID',($htmlsilver -bor $htmlbold),
		'Device Platform',($htmlsilver -bor $htmlbold),
		'Device Model',($htmlsilver -bor $htmlbold),
		'Device Managed',($htmlsilver -bor $htmlbold),
		'Device MAM',($htmlsilver -bor $htmlbold),
		'Device MDM',($htmlsilver -bor $htmlbold),
		'Device Name',($htmlsilver -bor $htmlbold),
		'Enrolled User',($htmlsilver -bor $htmlbold))
		$msg = "$count Devices configuration found"
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 ""
	}
}
#endregion

#region Notification Server
Function ProcessNSrv
{
	#Bypass certificate verification to enable access with XMS IP Address 
	[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

	#Connect to XMS server
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/authentication/login -Body $cred -Headers $headers -Method POST -Verbose:$false
	$headers.add("auth_token",$json.auth_token)

	#List Notification Server configuration
	$notif=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/notificationserver -Headers $headers -Method Get -Verbose:$false

	# Declare an array to collect our result objects
	$resultsarray =@()

	#$count will be the ‘loop counter’
	$count = $notif.list.length

	$Chapter++
	Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters Notification Server"

	Try
	{
		$Script:NotSrv = $notif
	}

	Catch
	{
		$Script:NotSrv = $Null
	}

	If(!$? -or $Null -eq $Script:NotSrv)
	{
		Write-Warning "No Notification Server were retrieved."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "Warning: No Notification Server were retrieved" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 0 "Warning: No Notification Server were retrieved"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Warning: No Notification Server were retrieved" "" $Null 0 $htmlbold
		}
		Return $False
	}
	ElseIf($? -and $Null -eq $Script:NotSrv)
	{
		Write-Warning "Notification Server retrieval was successful but no Notification Server were returned."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "Notification Server retrieval was successful but no Notification Server were returned." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 0 "Notification Server retrieval was successful but no Notification Server were returned."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Notification Server retrieval was successful but no Notification Server were returned." "" $Null 0 $htmlbold
		}
		Return $False
	}
	Else
	{
		If($notif -is [array])
		{
			[int]$Script:NumNotSrv = $count
		}
		Else
		{
			[int]$Script:NumNotSrv = 1
		}
		Write-Verbose "$(Get-Date): `t`t$($count) Notification Server configuration found"
		Return $True
	}
}

Function ProcessNotSrv
{
	If($MSWORD -or $PDF)
	{
		$Script:Selection.InsertNewPage()
		Write-Verbose "$(Get-Date): `t`t`tProcessing Notification Server information" 
		WriteWordLine 1 0 "Notification Server"
	}
	ElseIf($Text)
	{
		Write-Verbose "$(Get-Date): `t`t`tProcessing Notification Server information" 
		Line 0 ""
		Line 0 "Notification Server"
	}
	ElseIf($HTML)
	{
		Write-Verbose "$(Get-Date): `t`t`tProcessing Notification Server	information"
		WriteHTMLLine 0 0 ""
		WriteHTMLLine 1 0 "Notification Server"
		WriteHTMLLine 0 0 ""
	}
	
	OutputNotSrv $Script:NotSrv $Script:NumNotSrv
	If($MSWORD -or $PDF)
	{
		OutputNotSrvNoInternalGridLines $Script:NotSrv $Script:NumNotSrv
	}
}

Function OutputNotSrv
{
	#Bypass certificate verification to enable access with XMS IP Address
	[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

	#Connect to XMS server 
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/authentication/login -Body $cred -Headers $headers -Method POST -Verbose:$false
	$headers.add("auth_token",$json.auth_token)

	#List Notification Server Configuration
	$notif=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/notificationserver -Headers $headers -Method Get -Verbose:$false

	#Declare an array to collect our result objects
	$resultsarray =@()

	#$count will be the ‘loop counter’
	$count = $notif.list.length

	Param([object]$notif, [int]$count)
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 ""
		WriteWordLine 0 1 "$count Notification Server found on the XMS Server $XMS"
		WriteWordLine 0 1 ""

		## IB - replacement Services table generation utilising AddWordTable function

		## Create an array of hashtables to store our services
		[System.Collections.Hashtable[]] $NotSrvWordTable = @();
		## Create an array of hashtables to store references of cells that we wish to highlight after the table has been added
		[System.Collections.Hashtable[]] $HighlightedCells = @();
		[int] $CurrentNotSrvIndex = 0;
	}
	
	ElseIf($Text)
	{
		Line 0 ""
		Line 0 "`t$count Notification Server configuration found on the XMS Server $XMS"
		Line 0 ""
	}
	
	ElseIf($HTML)
	{
		$rowdata = @()
	}

	$CurrentNotSrvIndex=0
	for ($v=0;$v -lt $count; $v++)
	{
		ForEach($list in $notif) 
		{
			#Write-Verbose "$(Get-Date): `t`t`t Processing Devices $($Service.DisplayName)";

			If($MSWord -or $PDF)
			{
				## Add the required key/values to the hashtable
				$WordTableRowHash = @{ name = $list.list[$CurrentNotSrvIndex].name;
									   id = $list.list[$CurrentNotSrvIndex].id;
									   active = $list.list[$CurrentNotSrvIndex].active;
									   server = $list.list[$CurrentNotSrvIndex].server;
									   servertype = $list.list[$CurrentNotSrvIndex].servertype;
									 }

				## Add the hash to the array
				$NotSrvWordTable += $WordTableRowHash;
				$CurrentNotSrvIndex++
			}
			ElseIf($Text)
			{
				Line 0 "Name`t`t: " $list.list.name[$CurrentNotSrvIndex]
				Line 0 "ID`t`t: " $list.list.id[$CurrentNotSrvIndex]
				Line 0 "Active`t`t: " $list.list.active[$CurrentNotSrvIndex]
				Line 0 "Server`t`t: " $list.list.server[$CurrentNotSrvIndex]
				Line 0 "Server Type`t: " $list.list.servertype[$CurrentNotSrvIndex]
				Line 0 ""
				$CurrentNotSrvIndex++
			}
			ElseIf($HTML)
			{
				$HighlightedCells = $htmlwhite
				$rowdata += @(,(
				$list.list.name[$CurrentNotSrvIndex],$htmlwhite,
				$list.list.id[$CurrentNotSrvIndex],$htmlwhite,
				$list.list.active[$CurrentNotSrvIndex],$htmlwhite,
				$list.list.server[$CurrentNotSrvIndex],$htmlwhite,
				$list.list.servertype[$CurrentNotSrvIndex],$htmlwhite))
				$CurrentNotSrvIndex++
			}	
		}
	}

	If($MSWord -or $PDF)
	{
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		$Table = AddWordTable -Hashtable $NotSrvWordTable `
		-Columns name, id, active, server, servertype `
		-Headers "Name", "ID", "Active", "Server", "Server Type" `
		-Format -155 `
		
		## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
		SetWordCellFormat -Collection $Table -Size 10
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Color $wdColorWhite -Bold -BackgroundColor $wdColorRoyalBlue;
		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 40;
		$Table.Columns.Item(3).Width = 50;
		$Table.Columns.Item(4).Width = 250;
		$Table.Columns.Item(5).Width = 100;
		## IB - Set the required highlighted cells
		SetWordCellFormat -Coordinates $HighlightedCells -Table $Table -Bold -BackgroundColor $wdColorRed -Solid;

		#indent the entire table 1 tab stop
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)
		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($Text)
	{
	}
	ElseIf($HTML)
	{
		$columnHeaders = @(
		'Name',($htmlsilver -bor $htmlbold),
		'ID',($htmlsilver -bor $htmlbold),
		'Active',($htmlsilver -bor $htmlbold),
		'Server',($htmlsilver -bor $htmlbold),
		'Server Type',($htmlsilver -bor $htmlbold))
		$msg = "$count Notification Server configuration found"
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 ""
	}
}
#endregion

#region UsersGroups
Function ProcessUG
{
	#Bypass certificate verification to enable access with XMS IP Address 
	[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

	#Connect to XMS server
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/authentication/login -Body $cred -Headers $headers -Method POST -Verbose:$false
	$headers.add("auth_token",$json.auth_token)

	#List User Groups configuration
	$user=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/groups -Headers $headers -Method Get -Verbose:$false

	# Declare an array to collect our result objects
	$resultsarray =@()

	$Chapter++
	Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters Users Groups"

	Try
	{
		$Script:UG = $user
	}

	Catch
	{
		$Script:UG = $Null
	}

	If(!$? -or $Null -eq $Script:UG)
	{
		Write-Warning "No Users Groups were retrieved."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "Warning: No Users Groups were retrieved" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 0 "Warning: No Users Groups were retrieved"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Warning: No Users Groups were retrieved" "" $Null 0 $htmlbold
		}
		Return $False
	}
	ElseIf($? -and $Null -eq $Script:LUG)
	{
		Write-Warning "Users Groups retrieval was successful but no Users Groups were returned."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "Users Groups retrieval was successful but no Users Groups were returned." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 0 "Users Groups retrieval was successful but no Users Groups were returned."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Users Groups retrieval was successful but no Users Groups were returned." "" $Null 0 $htmlbold
		}
		Return $False
	}
	Else
	{
		If($user -is [array])
		{
			[int]$Script:NumLUG = $count
		}
		Else
		{
			[int]$Script:NumLUG = 1
		}
		Write-Verbose "$(Get-Date): `t`t$($count) Users Groups configuration found"
		Return $True
	}
}

Function ProcessUsersG
{
	If($MSWORD -or $PDF)
	{
		$Script:Selection.InsertNewPage()
		Write-Verbose "$(Get-Date): `t`t`tProcessing Users Groups information"
		WriteWordLine 1 0 "Users Groups"
	}
	ElseIf($Text)
	{
		Write-Verbose "$(Get-Date): `t`t`tProcessing Users Groups information"
		Line 0 ""
		Line 0 "Users Groups"
	}
	ElseIf($HTML)
	{
		Write-Verbose "$(Get-Date): `t`t`tProcessing Users Groups information"
		WriteHTMLLine 0 0 ""
		WriteHTMLLine 1 0 "Users Groups"
		WriteHTMLLine 0 0 ""
	}
	
	OutputUsersG $Script:UG $Script:NumUG
	If($MSWORD -or $PDF)
	{
		OutputUsersGNoInternalGridLines $Script:UG $Script:NumUG
	}
}

Function OutputUsersG
{
	#Bypass certificate verification to enable access with XMS IP Address
	[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

	#Connect to XMS server
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/authentication/login -Body $cred -Headers $headers -Method POST -Verbose:$false
	$headers.add("auth_token",$json.auth_token)

	#List Users Groups Configuration 
	$user=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/groups -Headers $headers -Method Get -Verbose:$false

	# Declare an array to collect our result objects
	$resultsarray =@()

	#$count will be the ‘loop counter’
	$count = $user.userGroups.length

	Param([object]$user, [int]$count)
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 ""
		WriteWordLine 0 1 "$count Users Groups found on the XMS Server $XMS"
		WriteWordLine 0 1 ""

		## IB - replacement Services table generation utilising AddWordTable function

		## Create an array of hashtables to store our services
		[System.Collections.Hashtable[]] $LUGWordTable = @();
		## Create an array of hashtables to store references of cells that we wish to highlight after the table has been added
		[System.Collections.Hashtable[]] $HighlightedCells = @();
		[int] $CurrentLUGIndex = 0;
	}
	
	ElseIf($Text)
	{
		Line 0 ""
		Line 0 "`t$count Users Groups configuration found on the XMS Server $XMS"
		Line 0 ""
	}
	
	ElseIf($HTML)
	{
		$rowdata = @()
	}

	$CurrentUGIndex=0
	for ($v=0;$v -lt $count; $v++)
	{
		ForEach($userGroups in $user) 
		{
			If($MSWord -or $PDF)
			{
				## Add the required key/values to the hashtable
				$WordTableRowHash = @{ name = $user.userGroups[$CurrentUGIndex].name;
									   uniqueName = $user.userGroups[$CurrentUGIndex].uniqueName;
									   uniqueId = $user.userGroups[$CurrentUGIndex].uniqueId;
									   domainname = $user.userGroups[$CurrentUGIndex].domainname;
									 }
				## Add the hash to the array
				$UGWordTable += $WordTableRowHash;
				$CurrentUGIndex++
			}
			ElseIf($Text)
			{
				Line 0 "Name`t`t`t: " $user.userGroups[$CurrentUGIndex].name
				Line 0 "Unique Name`t`t: " $user.userGroups[$CurrentUGIndex].uniqueName
				Line 0 "Unique Id`t`t: " $user.userGroups[$CurrentUGIndex].uniqueId
				Line 0 "Domain Name`t`t: " $user.userGroups[$CurrentUGIndex].domainname
				Line 0 ""
				$CurrentUGIndex++
			}
			ElseIf($HTML)
			{
				$HighlightedCells = $htmlwhite
				$rowdata += @(,(
				$user.userGroups[$CurrentUGIndex].name,$htmlwhite,
				$user.userGroups[$CurrentUGIndex].uniqueName,$htmlwhite,
				$user.userGroups[$CurrentUGIndex].uniqueId,$htmlwhite,
				$user.userGroups[$CurrentUGIndex].domainname,$htmlwhite))
				$CurrentUGIndex++
			}	
		}
	}

	If($MSWord -or $PDF)
	{
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		$Table = AddWordTable -Hashtable $UGWordTable `
		-Columns Name, uniqueName, UniqueId, domainname `
		-Headers "Name", "Unique Name", "Unique Id", "Domain Name" `
		-Format -155 `
		
		## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
		SetWordCellFormat -Collection $Table -Size 10
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Color $wdColorWhite -Bold -BackgroundColor $wdColorRoyalBlue;
		# Set Columns Width
		$Table.Columns.Item(1).Width = 220;
		$Table.Columns.Item(2).Width = 160;
		$Table.Columns.Item(3).Width = 160;
		$Table.Columns.Item(4).Width = 100;
		
		## IB - Set the required highlighted cells
		SetWordCellFormat -Coordinates $HighlightedCells -Table $Table -Bold -BackgroundColor $wdColorRed -Solid;

		#indent the entire table 1 tab stop
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)
		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($Text)
	{
	}
	ElseIf($HTML)
	{
		$columnHeaders = @(
		'Name',($htmlsilver -bor $htmlbold),
		'Unique Name',($htmlsilver -bor $htmlbold),
		'Unique Id',($htmlsilver -bor $htmlbold),
		'Domain Name',($htmlsilver -bor $htmlbold))
		$msg = "$count Users Groups configuration found"
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 ""
	}
}
#endregion

#region Delivery Groups
Function ProcessDgroups
{
	#Bypass certificate verification to enable access with XMS IP Address
	[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

	#Connect to XMS server
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/authentication/login -Body $cred -Headers $headers -Method POST -Verbose:$false
	$headers.add("auth_token",$json.auth_token)

	#List Delivery Groups configuration
	$dgBody=
	'
	{
	"start": "0",
	"limit": "1000"
	}
	'
	$dgroup=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/deliverygroups/filter" -Body $dgBody -Headers $headers -Method Post -Verbose:$false

	# Declare an array to collect our result objects
	$resultsarray =@()

	#$count will be the ‘loop counter’
	$count = $dgroup.dglistdata.dglist.length

	$Chapter++
	Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters Delivery Groups"

	Try
	{
		$Script:dgroup = $dgroup
	}

	Catch
	{
		$Script:dgroup = $Null
	}

	If(!$? -or $Null -eq $Script:dgroup)
	{
		Write-Warning "No Delivery Groups were retrieved."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "Warning: No Delivery Groups were retrieved" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 0 "Warning: No Delivery Groups were retrieved"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Warning: No Delivery Groups were retrieved" "" $Null 0 $htmlbold
		}
		Return $False
	}
	ElseIf($? -and $Null -eq $Script:dgroup)
	{
		Write-Warning "Delivery Groups retrieval was successful but no Delivery Groups were returned."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "Delivery Groups retrieval was successful but no Delivery Groups were returned." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 0 "Delivery Groups retrieval was successful but no Delivery Groups were returned."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Delivery Groups retrieval was successful but no Delivery Groups were returned." "" $Null 0 $htmlbold
		}
		Return $False
	}
	Else
	{
		If($dgroup -is [array])
		{
			[int]$Script:NumDGROUP = $count
		}
		Else
		{
			[int]$Script:NumDGROUP = 1
		}
		Write-Verbose "$(Get-Date): `t`t$($count) Delivery Groups configuration found"
		Return $True
	}
}

Function ProcessDelGroups
{
	If($MSWORD -or $PDF)
	{
		$Script:Selection.InsertNewPage()
		Write-Verbose "$(Get-Date): `t`t`tProcessing Delivery Groups information"
		WriteWordLine 1 0 "Delivery Groups"
	}
	ElseIf($Text)
	{
		Write-Verbose "$(Get-Date): `t`t`tProcessing Delivery Groups information"
		Line 0 ""
		Line 0 "Delivery Groups"
	}
	ElseIf($HTML)
	{
		Write-Verbose "$(Get-Date): `t`t`tProcessing Delivery Groups information"
		WriteHTMLLine 0 0 ""
		WriteHTMLLine 1 0 "Delivery Groups"
		WriteHTMLLine 0 0 ""
	}
	
	OutputDelGroups $Script:dgroup $Script:NumDGROUP
	If($MSWORD -or $PDF)
	{
		OutputDelGroupsNoInternalGridLines $Script:dgroup $Script:NumDGROUP
	}
}

Function Table
{
	ForEach($dglistdata in $dgroup) 
	{#$CurrentDgroupIndex=0
		$count = $dgroup.dglistdata.dglist.length
		for ($v=0;$v -lt $count; $v++)
		{
			If($MSWord -or $PDF)
			{
				## Add the required key/values to the hashtable
				$WordTableRowHash = @{ id = $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].id;
									   name = $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].name;
									   description = $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].description;
									   disabled = $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].disabled;
									   nbsuccess = $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].nbsuccess;
									   nbfailure = $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].nbfailure;
									   nbpending = $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].nbpending;
									   enrollmentprofilename = $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].enrollmentprofilename;
									 }
				$DgroupWordTable += $WordTableRowHash;					 
				$CurrentDgroupIndex++
			}
			ElseIf($Text)
			{
				Line 0 "ID`t`t`t`t: " $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].id
				Line 0 "Name`t`t`t`t: " $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].name
				Line 0 "Description`t`t`t: " $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].description
				Line 0 "Disabled`t`t`t: " $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].disabled
				Line 0 "Application Name`t`t: " $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].applications.name
				Line 0 "Application Required`t`t: " $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].applications.required
				Line 0 "Device Policies Name`t`t: " $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].devicepolicies.name
				Line 0 "Nb of Success`t`t`t: " $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].nbsuccess
				Line 0 "Nb of Failure`t`t`t: " $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].nbfailure
				Line 0 "Nb of Pending`t`t`t: " $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].nbpending
				Line 0 "Enrollment Profile Name`t`t: " $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].enrollmentprofilename
				Line 0 ""
				$CurrentDgroupIndex++
			}
			ElseIf($HTML) 
			{
				$HighlightedCells = $htmlwhite
				$rowdata += @(,(
				$dglistdata.dglistdata.dglist[$CurrentDgroupIndex].id,$htmlwhite,
				$dglistdata.dglistdata.dglist[$CurrentDgroupIndex].name,$htmlwhite,
				$dglistdata.dglistdata.dglist[$CurrentDgroupIndex].description,$htmlwhite,
				$dglistdata.dglistdata.dglist[$CurrentDgroupIndex].disabled,$htmlwhite,
				$dglistdata.dglistdata.dglist[$CurrentDgroupIndex].applications.name,$htmlwhite,
				$dglistdata.dglistdata.dglist[$CurrentDgroupIndex].applications.required,$htmlwhite,
				$dglistdata.dglistdata.dglist[$CurrentDgroupIndex].devicepolicies.name,$htmlwhite,
				$dglistdata.dglistdata.dglist[$CurrentDgroupIndex].nbsuccess,$htmlwhite,
				$dglistdata.dglistdata.dglist[$CurrentDgroupIndex].nbfailure,$htmlwhite,
				$dglistdata.dglistdata.dglist[$CurrentDgroupIndex].nbpending,$htmlwhite,
				$dglistdata.dglistdata.dglist[$CurrentDgroupIndex].enrollmentprofilename,$htmlwhite))
				$CurrentDgroupIndex++
			}	
		}
	
		If($MSWord -or $PDF)
		{
			## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
			$Table = AddWordTable -Hashtable $DgroupWordTable `
			-Columns id, name, description, disabled, nbsuccess, nbfailure, nbpending, enrollmentprofilename `
			-Headers "ID", "Name", "Description", "Disabled", "Nb of Success", "Nb of Failure", "Nb of Pending", "Enrollment Profile Name" `
			-Format -155 `
			
			## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
			SetWordCellFormat -Collection $Table -Size 10
			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Color $wdColorWhite -Bold -BackgroundColor $wdColorRoyalBlue;
			# Set Columns Width
			$Table.Columns.Item(1).Width = 30;
			$Table.Columns.Item(2).Width = 125;
			$Table.Columns.Item(3).Width = 125;
			$Table.Columns.Item(4).Width = 70;
			$Table.Columns.Item(5).Width = 60;
			$Table.Columns.Item(6).Width = 60;
			$Table.Columns.Item(7).Width = 60;
			$Table.Columns.Item(8).Width = 110;
			## IB - Set the required highlighted cells
			SetWordCellFormat -Coordinates $HighlightedCells -Table $Table -Bold -BackgroundColor $wdColorRed -Solid;
			#indent the entire table 1 tab stop
			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)
			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 2 ""
					
		}
		ElseIf($Text)
		{
		}
		ElseIf($HTML)
		{
			$columnHeaders = @(
			'ID',($htmlsilver -bor $htmlbold),
			'Name',($htmlsilver -bor $htmlbold),
			'Description',($htmlsilver -bor $htmlbold),
			'Disabled',($htmlsilver -bor $htmlbold),
			'Application Name',($htmlsilver -bor $htmlbold),
			'Application Required',($htmlsilver -bor $htmlbold),
			'Device Policies',($htmlsilver -bor $htmlbold),
			'Nb of Success',($htmlsilver -bor $htmlbold),
			'Nb of Failure',($htmlsilver -bor $htmlbold),
			'Nb of Pending',($htmlsilver -bor $htmlbold),
			'Enrollment Profile Name',($htmlsilver -bor $htmlbold))
			$msg = "$count Delivery Groups configuration found"
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
			WriteHTMLLine 0 0 ""
		}
	}
}

Function OutputDelGroups
{
	#Bypass certificate verification to enable access with XMS IP Address
	[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

	#Connect to XMS server
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/authentication/login -Body $cred -Headers $headers -Method POST -Verbose:$false
	$headers.add("auth_token",$json.auth_token)

	#List Local Users Groups Configuration
	$dgBody=
	'
	{
	"start": "0",
	"limit": "1000"
	}
	'
	$dgroup=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/deliverygroups/filter" -Body $dgBody -Headers $headers -Method Post -Verbose:$false

	# Declare an array to collect our result objects
	$resultsarray =@()

	#$count will be the ‘loop counter’
	$count = $dgroup.dglistdata.dglist.length

	Param([object]$dgroup, [int]$count)
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 ""
		WriteWordLine 0 1 "$count Delivery Groups found on the XMS Server $XMS"
		WriteWordLine 0 1 ""
		$CurrentDgroupIndex=0
		For ($v=0;$v -lt $count; $v++)
		{
			$nbapp = $dgroup.dglistdata.dglist[$CurrentDgroupIndex].applications.length
			$nbpol = $dgroup.dglistdata.dglist[$CurrentDgroupIndex].devicePolicies.length
			ForEach($dglistdata in $dgroup) 
			{
				WriteWordLine 2 0 $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].name 
				WriteWordLine 0 1 ""
				Table
		
				WriteWordLine 0 1 "Applications:" "" "Calibri" 10 $false $true
				WriteWordLine 0 1 ""
				for ($w=0; $w -lt $nbapp; $w++)
				{
					WriteWordLine 0 2 $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].applications[$w].name "" "Calibri" 10 $false $false
					WriteWordLine 0 4 "Required: "$dglistdata.dglistdata.dglist[$CurrentDgroupIndex].applications[$w].required $null 10 $true $true
				}
				WriteWordLine 0 1 ""
				WriteWordLine 0 1 "Device Policies:" "" "Calibri" 10 $false $true
				WriteWordLine 0 1 ""
				for ($x=0; $x -lt $nbpol; $x++)
				{
					WriteWordLine 0 2 $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].devicePolicies[$x].name "" "Calibri" 10 $false $false
				}
				$Script:Selection.InsertNewPage()
				$CurrentDgroupIndex++
			}
		}
		## IB - replacement Services table generation utilising AddWordTable function

		## Create an array of hashtables to store our services
		[System.Collections.Hashtable[]] $DgroupWordTable = @();
		## Create an array of hashtables to store references of cells that we wish to highlight after the table has been added
		[System.Collections.Hashtable[]] $HighlightedCells = @();
		[int] $CurrentDgroupIndex = 0;
	}
	ElseIf($Text)
	{
		Line 0 ""
		Line 0 "`t$count Delivery Groups configuration found on the XMS Server $XMS"
		Line 0 ""
	}
	ElseIf($HTML)
	{
		$rowdata = @()
	}

	If($MSWord -or $PDF)
			{
			$CurrentDgroupIndex=0
			For ($v=0;$v -lt $count; $v++)
			{
			$nbapp = $dgroup.dglistdata.dglist[$CurrentDgroupIndex].applications.length
			$nbpol = $dgroup.dglistdata.dglist[$CurrentDgroupIndex].devicePolicies.length
			ForEach($dglistdata in $dgroup) 
			{
				## Add the required key/values to the hashtable
				$WordTableRowHash = @{ id = $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].id;
									   name = $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].name;
									   description = $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].description;
									   disabled = $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].disabled;
									 } 
									 $DgroupWordTable += $WordTableRowHash;
									 for ($w=0; $w -lt $nbapp; $w++)
									 {
									 $WordTableRowHash = @{
  									   appname = $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].applications[$w].name;
									   appreq = $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].applications[$w].required;
									 }
									 }
									 $DgroupWordTable += $WordTableRowHash;
									 for ($x=0; $x -lt $nbpol; $x++)
									 {
									 $WordTableRowHash = @{
									   devpol = $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].devicePolicies[$x].name;
									 }
									 }
									 $DgroupWordTable += $WordTableRowHash;
				$WordTableRowHash = @{ 
									   nbsuccess = $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].nbsuccess;
									   nbfailure = $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].nbfailure;
									   nbpending = $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].nbpending;
									   enrollmentprofilename = $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].enrollmentprofilename;
									 }

				## Add the hash to the array
				$DgroupWordTable += $WordTableRowHash;
				$CurrentDgroupIndex++
			}
			}
			}
			ElseIf($Text)
			{
			$CurrentDgroupIndex=0
			For ($v=0;$v -lt $count; $v++)
			{
			$nbapp = $dgroup.dglistdata.dglist[$CurrentDgroupIndex].applications.length
			$nbpol = $dgroup.dglistdata.dglist[$CurrentDgroupIndex].devicePolicies.length
			ForEach($dglistdata in $dgroup) 
			{
				Line 0 "ID`t`t`t`t: " $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].id
				Line 0 "Name`t`t`t`t: " $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].name
				Line 0 "Description`t`t`t: " $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].description
				Line 0 "Disabled`t`t`t: " $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].disabled
				for ($w=0; $w -lt $nbapp; $w++)
				{
				Line 0 "Application Name`t`t: " $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].applications[$w].name
				Line 0 "Application Required`t`t: " $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].applications[$w].required
				}
				for ($x=0; $x -lt $nbpol; $x++)
				{
				Line 0 "Device Policies Name`t`t: " $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].devicepolicies[$x].name
				}
				Line 0 "Nb of Success`t`t`t: " $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].nbsuccess
				Line 0 "Nb of Failure`t`t`t: " $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].nbfailure
				Line 0 "Nb of Pending`t`t`t: " $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].nbpending
				Line 0 "Enrollment Profile Name`t`t: " $dglistdata.dglistdata.dglist[$CurrentDgroupIndex].enrollmentprofilename
				Line 0 ""
				$CurrentDgroupIndex++
			}
			}
			}
			ElseIf($HTML) 
			{
			$CurrentDgroupIndex=0
			For ($v=0;$v -lt $count; $v++)
			{
			$nbapp = $dgroup.dglistdata.dglist[$CurrentDgroupIndex].applications.length
			$nbpol = $dgroup.dglistdata.dglist[$CurrentDgroupIndex].devicePolicies.length
			ForEach($dglistdata in $dgroup) 
			{
				$HighlightedCells = $htmlwhite
				$rowdata += @(,(
				$dglistdata.dglistdata.dglist[$CurrentDgroupIndex].id,$htmlwhite,
				$dglistdata.dglistdata.dglist[$CurrentDgroupIndex].name,$htmlwhite,
				$dglistdata.dglistdata.dglist[$CurrentDgroupIndex].description,$htmlwhite,
				$dglistdata.dglistdata.dglist[$CurrentDgroupIndex].disabled,$htmlwhite)
				for ($w=0; $w -lt $nbapp; $w++)
				{
				$rowdata += @(
				$dglistdata.dglistdata.dglist[$CurrentDgroupIndex].applications[$w].name,$htmlwhite,
				$dglistdata.dglistdata.dglist[$CurrentDgroupIndex].applications[$w].required,$htmlwhite)
				}
				for ($x=0; $x -lt $nbpol; $x++)
				{
				$rowdata += @(
				$dglistdata.dglistdata.dglist[$CurrentDgroupIndex].devicepolicies[$x].name,$htmlwhite)
				}
				$rowdata += @(
				$dglistdata.dglistdata.dglist[$CurrentDgroupIndex].nbsuccess,$htmlwhite,
				$dglistdata.dglistdata.dglist[$CurrentDgroupIndex].nbfailure,$htmlwhite,
				$dglistdata.dglistdata.dglist[$CurrentDgroupIndex].nbpending,$htmlwhite,
				$dglistdata.dglistdata.dglist[$CurrentDgroupIndex].enrollmentprofilename,$htmlwhite))
				$CurrentDgroupIndex++
			}
			}
						
		}
	
	
	If($MSWord -or $PDF)
	{

		$Table = $Null
	}
	ElseIf($Text)
	{
	}
	ElseIf($HTML)
	{
		$columnHeaders = @(
		'ID',($htmlsilver -bor $htmlbold),
		'Name',($htmlsilver -bor $htmlbold),
		'Description',($htmlsilver -bor $htmlbold),
		'Disabled',($htmlsilver -bor $htmlbold),
		'Application Name',($htmlsilver -bor $htmlbold),
		'Required',($htmlsilver -bor $htmlbold),
		'Device Policies',($htmlsilver -bor $htmlbold),
		'Nb of Success',($htmlsilver -bor $htmlbold),
		'Nb of Failure',($htmlsilver -bor $htmlbold),
		'Nb of Pending',($htmlsilver -bor $htmlbold),
		'Enrollment Profile Name',($htmlsilver -bor $htmlbold))
		$msg = "$count Delivery Groups configuration found"
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 ""
	}
}
#endregion

#region Enrollment Mode
Function ProcessEMode
{
	#Bypass certificate verification to enable access with XMS IP Address
	[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

	# Connect to XMS server
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/authentication/login -Body $cred -Headers $headers -Method POST -Verbose:$false
	$headers.add("auth_token",$json.auth_token)

	# List Enrollment Mode configuration
	$enrollment=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/enrollment/modes" -Headers $headers -Method Get -Verbose:$false

	# Declare an array to collect our result objects
	$resultsarray =@()

	#$count will be the ‘loop counter’
	$count = $enrollment.enrollmentmodes.enrollmentmodes.length

	$Chapter++
	Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters Enrollment Mode"

	Try
	{
		$Script:emode = $enrollment
	}

	Catch
	{
		$Script:emode = $Null
	}

	If(!$? -or $Null -eq $Script:emode)
	{
		Write-Warning "No Enrollment Mode were retrieved."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "Warning: No Enrollment Mode were retrieved" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 0 "Warning: No Enrollment Mode were retrieved"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Warning: No Enrollment Mode were retrieved" "" $Null 0 $htmlbold
		}
		Return $False
	}
	ElseIf($? -and $Null -eq $Script:emode)
	{
		Write-Warning "Enrollment Mode retrieval was successful but no Enrollment Mode were returned."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "Enrollment Mode retrieval was successful but no Enrollment Mode were returned." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 0 "Enrollment Mode retrieval was successful but no Enrollment Mode were returned."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Enrollment Mode retrieval was successful but no Enrollment Mode were returned." "" $Null 0 $htmlbold
		}
		Return $False
	}
	Else
	{
		If($emode -is [array])
		{
			[int]$Script:Numemode = $count
		}
		Else
		{
			[int]$Script:Numemode = 1
		}
		Write-Verbose "$(Get-Date): `t`t$($count) Enrollment Mode configuration found"
		Return $True
	}
}

Function ProcessEnrollMode
{
	If($MSWORD -or $PDF)
	{
		$Script:Selection.InsertNewPage()
		Write-Verbose "$(Get-Date): `t`t`tProcessing Enrollment Mode information" 
		WriteWordLine 1 0 "Enrollment Mode"
	}
	ElseIf($Text)
	{
		Write-Verbose "$(Get-Date): `t`t`tProcessing Enrollment Mode information"
		Line 0 ""
		Line 0 "Enrollment Mode"
	}
	ElseIf($HTML)
	{
		Write-Verbose "$(Get-Date): `t`t`tProcessing Enrollment Mode information"
		WriteHTMLLine 0 0 ""
		WriteHTMLLine 1 0 "Enrollment Mode"
		WriteHTMLLine 0 0 ""
	}
	
	OutputEnrollMode $Script:emode $Script:Numemode
	If($MSWORD -or $PDF)
	{
		OutputEnrollModeNoInternalGridLines $Script:emode $Script:Numemode
	}
}

Function OutputEnrollMode
{
	#Bypass certificate verification to enable access with XMS IP Address
	[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
	
	#Connect to XMS server
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/authentication/login -Body $cred -Headers $headers -Method POST -Verbose:$false
	$headers.add("auth_token",$json.auth_token)

	#List Enrollment Mode Configuration
	$enrollment=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/enrollment/modes"  -Headers $headers -Method Get -Verbose:$false

	#Declare an array to collect our result objects
	$resultsarray =@()

	#$count will be the ‘loop counter’
	$count = $enrollment.enrollmentmodes.enrollmentmodes.length

	Param([object]$enrollment, [int]$count)
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 ""
		WriteWordLine 0 1 "$count Enrollment Mode found on the XMS Server $XMS"
		WriteWordLine 0 1 ""

		## IB - replacement Services table generation utilising AddWordTable function

		## Create an array of hashtables to store our services
		[System.Collections.Hashtable[]] $emodeWordTable = @();
		## Create an array of hashtables to store references of cells that we wish to highlight after the table has been added
		[System.Collections.Hashtable[]] $HighlightedCells = @();
		[int] $CurrentemodeIndex = 0;
	}
	
	ElseIf($Text)
	{
		Line 0 ""
		Line 0 "`t$count Enrollment Mode configuration found on the XMS Server $XMS"
		Line 0 ""
	}
	
	ElseIf($HTML)
	{
		$rowdata = @()
	}

	$CurrentemodeIndex=0
	ForEach($enrollmentmodes in $enrollment) 
	{
		for ($v=0;$v -lt $count; $v++)
		{
			#Write-Verbose "$(Get-Date): `t`t`t Processing Enrollment Mode $($Service.DisplayName)";

			If($MSWord -or $PDF)
			{
				## Add the required key/values to the hashtable
				$WordTableRowHash = @{ name = $enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].name;
									   modedisplayname = $enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].modedisplayname;
									   validdurationmillis = $enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].validdurationmillis;
									   maxtry = $enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].maxtry;
									   enabled = $enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].enabled;
									   shpmode = $enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].shpmode;
									   defaultable = $enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].defaultable;
									   requiringidentification = $enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].requiringidentification;
									   requiringauthentication = $enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].requiringauthentication;
									   requiringtoken = $enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].requiringtoken;
									   requiringsecret = $enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].requiringsecret;
									   requiringcertificate = $enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].requiringcertificate;
									   default = $enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].default;
									 }

				## Add the hash to the array
				$emodeWordTable += $WordTableRowHash;
				$CurrentemodeIndex++
			}
			ElseIf($Text)
			{
				Line 0 "Name`t`t`t`t: " $enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].name
				Line 0 "Display Name`t`t`t: " $enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].modedisplayname
				Line 0 "Validation Duration (ms)`t: " $enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].validdurationmillis
				Line 0 "Max Try`t`t`t`t: " $enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].maxtry
				Line 0 "Enabled`t`t`t`t: " $enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].enabled
				Line 0 "Self Help Portal`t`t: " $enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].shpmode
				Line 0 "Default Table`t`t`t: " $enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].defaultable
				Line 0 "Require Identification`t`t: " $enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].requiringidentification
				Line 0 "Require Authentication`t`t: " $enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].requiringauthentication
				Line 0 "Require Token`t`t`t: " $enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].requiringtoken
				Line 0 "Require Secret`t`t`t: " $enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].requiringsecret
				Line 0 "Require Certificate`t`t: " $enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].requiringcertificate
				Line 0 "Default`t`t`t`t: " $enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].default
				Line 0 ""
				$CurrentemodeIndex++
			}
			ElseIf($HTML) 
			{
				$HighlightedCells = $htmlwhite
				$rowdata += @(,(
				$enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].name,$htmlwhite,
				$enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].modedisplayname,$htmlwhite,
				$enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].validdurationmillis,$htmlwhite,
				$enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].maxtry,$htmlwhite,
				$enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].enabled,$htmlwhite,
				$enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].shpmode,$htmlwhite,
				$enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].defaultable,$htmlwhite,
				$enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].requiringidentification,$htmlwhite,
				$enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].requiringauthentication,$htmlwhite,
				$enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].requiringtoken,$htmlwhite,
				$enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].requiringsecret,$htmlwhite,
				$enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].requiringcertificate,$htmlwhite,
				$enrollmentmodes.enrollmentmodes.enrollmentmodes[$CurrentemodeIndex].default,$htmlwhite))
				$CurrentemodeIndex++
			}	
		}
	}

	If($MSWord -or $PDF)
	{
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		$Table = AddWordTable -Hashtable $emodeWordTable `
		-Columns name, modedisplayname, validdurationmillis, maxtry, enabled, shpmode, defaultable, requiringidentification, requiringauthentication, requiringtoken, requiringsecret, requiringcertificate, default `
		-Headers "Name", "Display Name", "Validation Duration (ms)", "Max Try", "Enabled", "Self Help Portal", "Default Table", "Require Identification", "Require Authentication", "Require Token", "Require Secret", "Require Certificate", "Default" `
		-Format -155 `
		
		## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
		## AP - Set the size for this table to 10
		SetWordCellFormat -Collection $Table -Size 10
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Color $wdColorWhite -Bold -BackgroundColor $wdColorRoyalBlue;
		# Set Columns Width
		$Table.Columns.Item(1).Width = 50;
		$Table.Columns.Item(2).Width = 50;
		$Table.Columns.Item(3).Width = 55;
		$Table.Columns.Item(4).Width = 45;
		$Table.Columns.Item(5).Width = 50;
		$Table.Columns.Item(6).Width = 50;
		$Table.Columns.Item(7).Width = 50;
		$Table.Columns.Item(8).Width = 50;
		$Table.Columns.Item(9).Width = 50;
		$Table.Columns.Item(10).Width = 50;
		$Table.Columns.Item(11).Width = 50;
		$Table.Columns.Item(12).Width = 55;
		$Table.Columns.Item(13).Width = 45;
		## IB - Set the required highlighted cells
		SetWordCellFormat -Coordinates $HighlightedCells -Table $Table -Bold -BackgroundColor $wdColorRed -Solid;

		#indent the entire table 1 tab stop
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)
		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($Text)
	{
	}
	ElseIf($HTML)
	{
		$columnHeaders = @(
		'Name',($htmlsilver -bor $htmlbold),
		'Display Name',($htmlsilver -bor $htmlbold),
		'Validation Duration (ms)',($htmlsilver -bor $htmlbold),
		'Max Try',($htmlsilver -bor $htmlbold),
		'Enabled',($htmlsilver -bor $htmlbold),
		'Self Help Portal',($htmlsilver -bor $htmlbold),
		'Default Table',($htmlsilver -bor $htmlbold),
		'Require Identification',($htmlsilver -bor $htmlbold),
		'Require Authentication',($htmlsilver -bor $htmlbold),
		'Require Token',($htmlsilver -bor $htmlbold),
		'Require Secret',($htmlsilver -bor $htmlbold),
		'Require Certificate',($htmlsilver -bor $htmlbold),
		'Default',($htmlsilver -bor $htmlbold))
		$msg = "$count Enrollment Mode configuration found"
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 ""
	}
}
#endregion

#region RBAC
Function ProcessRBAC
{
	#Bypass certificate verification to enable access with XMS IP Address
	[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

	#Connect to XMS server 
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/authentication/login -Body $cred -Headers $headers -Method POST -Verbose:$false
	$headers.add("auth_token",$json.auth_token)

	#List RBAC configuration 
	$rbac=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/rbac/roles" -Headers $headers -Method Get -Verbose:$false

	# Declare an array to collect our result objects
	$resultsarray =@()

	#$count will be the ‘loop counter’
	$count = $rbac.roles.length
	$Chapter++
	Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters RBAC"

	Try
	{
		$Script:rbac = $rbac
	}

	Catch
	{
		$Script:rbac = $Null
	}

	If(!$? -or $Null -eq $Script:rbac)
	{
		Write-Warning "No RBAC were retrieved."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "Warning: No RBAC were retrieved" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 0 "Warning: No RBAC were retrieved"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Warning: No RBAC were retrieved" "" $Null 0 $htmlbold
		}
		Return $False
	}
	ElseIf($? -and $Null -eq $Script:rbac)
	{
		Write-Warning "RBAC retrieval was successful but no RBAC were returned."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "RBAC retrieval was successful but no RBAC were returned." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 0 "Enrollment RBAC was successful but no RBAC were returned."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "RBAC retrieval was successful but no RBAC were returned." "" $Null 0 $htmlbold
		}
		Return $False
	}
	Else
	{
		If($rbac -is [array])
		{
			[int]$Script:Numrbac = $count
		}
		Else
		{
			[int]$Script:Numrbac = 1
		}
		Write-Verbose "$(Get-Date): `t`t$($count) RBAC configuration found"
		Return $True
	}
}

Function ProcessRBAControl
{
	If($MSWORD -or $PDF)
	{
		$Script:Selection.InsertNewPage()
		Write-Verbose "$(Get-Date): `t`t`tProcessing RBAC information"
		WriteWordLine 1 0 "Role-Based Access Control"
	}
	ElseIf($Text)
	{
		Write-Verbose "$(Get-Date): `t`t`tProcessing RBAC Mode information"
		Line 0 ""
		Line 0 "Role-Based Access Control"
	}
	ElseIf($HTML)
	{
		Write-Verbose "$(Get-Date): `t`t`tProcessing RBAC Mode information"
		WriteHTMLLine 0 0 ""
		WriteHTMLLine 1 0 "Role-Based Access Control"
		WriteHTMLLine 0 0 ""
	}
	
	OutputRBAControl $Script:rbac $Script:Numrbac
	If($MSWORD -or $PDF)
	{
		OutputRBAControlNoInternalGridLines $Script:rbac $Script:Numrbac
	}
}

Function OutputRBAControl
{
	#Bypass certificate verification to enable access with XMS IP Address
	[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

	# Connect to XMS server
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/authentication/login -Body $cred -Headers $headers -Method POST -Verbose:$false
	$headers.add("auth_token",$json.auth_token)

	#List Role-Based Access Control Configuration
	$rbac=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/rbac/roles" -Headers $headers -Method Get -Verbose:$false

	# Declare an array to collect our result objects
	$resultsarray =@()

	#$count will be the ‘loop counter'
	$count = $rbac.roles.length

	Param([object]$rbac, [int]$count)
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 ""
		WriteWordLine 0 1 "$count Role-Based Access Control found on the XMS Server $XMS"
		WriteWordLine 0 1 ""

		## IB - replacement Services table generation utilising AddWordTable function

		## Create an array of hashtables to store our services
		[System.Collections.Hashtable[]] $rbacWordTable = @();
		## Create an array of hashtables to store references of cells that we wish to highlight after the table has been added
		[System.Collections.Hashtable[]] $HighlightedCells = @();
		[int] $CurrentrbacIndex = 0;
	}
	
	ElseIf($Text)
	{
		Line 0 ""
		Line 0 "`t$count Role-Based Access Control configuration found on the XMS Server $XMS"
		Line 0 ""
	}
	
	ElseIf($HTML)
	{
		$rowdata = @()
	}

	$CurrentrbacIndex=0
	ForEach($roles in $rbac) 
	{
		for ($v=0;$v -lt $count; $v++)
		{
			#Write-Verbose "$(Get-Date): `t`t`t Processing Enrollment Mode $($Service.DisplayName)";

			If($MSWord -or $PDF)
			{
				## Add the required key/values to the hashtable
				$WordTableRowHash = @{ rolename = $rbac.roles[$CurrentrbacIndex];
								 }
				## Add the hash to the array
				$rbacWordTable += $WordTableRowHash;
				$CurrentrbacIndex++
			}
			ElseIf($Text)
			{
				Line 0 "Role Name`t`t: " $rbac.roles[$CurrentrbacIndex]
				Line 0 ""
				$CurrentrbacIndex++
			}
			ElseIf($HTML) 
			{
				$rowdata += @(,($rbac.roles[$CurrentrbacIndex],$htmlwhite))
				$CurrentrbacIndex++
			}	
		}
	}

	If($MSWord -or $PDF)
	{
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		$Table = AddWordTable -Hashtable $rbacWordTable `
		-Columns rolename `
		-Headers "Role Name" `
		-Format -155 `
		
		## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
		## AP - Set the size for this table to 8
		SetWordCellFormat -Collection $Table -Size 10
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Color $wdColorWhite -Bold -BackgroundColor $wdColorRoyalBlue;
		$Table.Columns.Item(13).Width = 640;
		## IB - Set the required highlighted cells
		SetWordCellFormat -Coordinates $HighlightedCells -Table $Table -Bold -BackgroundColor $wdColorRed -Solid;

		#indent the entire table 1 tab stop
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)
		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($Text)
	{
	}
	ElseIf($HTML)
	{
		$columnHeaders = @('Role Name',($htmlsilver -bor $htmlbold))
		$msg = "$count Role-Based Access Control configuration found"
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 ""
	}
}
#endregion

#region Client Properties
Function ProcessCP
{
	#Bypass certificate verification to enable access with XMS IP Address
	[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

	# Connect to XMS server
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/authentication/login -Body $cred -Headers $headers -Method POST -Verbose:$false
	$headers.add("auth_token",$json.auth_token)

	# List Client Properties configuration
	$clientprop=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/clientproperties" -Headers $headers -Method Get -Verbose:$false

	# Declare an array to collect our result objects
	$resultsarray =@()

	#$count will be the ‘loop counter’
	$count = $clientprop.allclientproperties.length

	$Chapter++
	Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters Client Properties"

	Try
	{
		$Script:cprop = $clientprop
	}

	Catch
	{
		$Script:cprop = $Null
	}

	If(!$? -or $Null -eq $Script:cprop)
	{
		Write-Warning "No Client Properties were retrieved."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "Warning: No Client Properties were retrieved" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 0 "Warning: No Client Properties were retrieved"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Warning: No Client Properties were retrieved" "" $Null 0 $htmlbold
		}
		Return $False
	}
	ElseIf($? -and $Null -eq $Script:cprop)
	{
		Write-Warning "Client Properties retrieval was successful but no Client Properties were returned."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "Client Properties retrieval was successful but no Client Properties were returned." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 0 "Enrollment Client Properties was successful but no Client Properties were returned."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Client Properties retrieval was successful but no Client Properties were returned." "" $Null 0 $htmlbold
		}
		Return $False
	}
	Else
	{
		If($clientprop -is [array])
		{
			[int]$Script:Numcprop = $count
		}
		Else
		{
			[int]$Script:Numcprop = 1
		}
		Write-Verbose "$(Get-Date): `t`t$($count) Client Properties configuration found"
		Return $True
	}
}

Function ProcessCprop
{
	If($MSWORD -or $PDF)
	{
		$Script:Selection.InsertNewPage()
		Write-Verbose "$(Get-Date): `t`t`tProcessing Client Properties information"
		WriteWordLine 1 0 "Client Properties"
	}
	ElseIf($Text)
	{
		Write-Verbose "$(Get-Date): `t`t`tProcessing Client Properties Mode information"
		Line 0 ""
		Line 0 "Client Properties"
	}
	ElseIf($HTML)
	{
		Write-Verbose "$(Get-Date): `t`t`tProcessing Client Properties Mode information"
		WriteHTMLLine 0 0 ""
		WriteHTMLLine 1 0 "Client Properties"
		WriteHTMLLine 0 0 ""
	}
	
	OutputCprop $Script:cprop $Script:Numcprop
	If($MSWORD -or $PDF)
	{
		OutputCpropNoInternalGridLines $Script:cprop $Script:Numcprop
	}
}

Function OutputCprop
{
	#Bypass certificate verification to enable access with XMS IP Address
	[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

	#Connect to XMS server
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/authentication/login -Body $cred -Headers $headers -Method POST -Verbose:$false
	$headers.add("auth_token",$json.auth_token)

	#List Client Properties Configuration
	$clientprop=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/clientproperties" -Headers $headers -Method Get -Verbose:$false

	# Declare an array to collect our result objects
	$resultsarray =@()

	#$count will be the ‘loop counter’
	$count = $clientprop.allclientproperties.length

	Param([object]$clientprop, [int]$count)
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 ""
		WriteWordLine 0 1 "$count Client Properties found on the XMS Server $XMS"
		WriteWordLine 0 1 ""

		## IB - replacement Services table generation utilising AddWordTable function

		## Create an array of hashtables to store our services
		[System.Collections.Hashtable[]] $cpropWordTable = @();
		## Create an array of hashtables to store references of cells that we wish to highlight after the table has been added
		[System.Collections.Hashtable[]] $HighlightedCells = @();
		[int] $CurrentcpropIndex = 0;
	}
	
	ElseIf($Text)
	{
		Line 0 ""
		Line 0 "`t$count Client Properties configuration found on the XMS Server $XMS"
		Line 0 ""
	}
	
	ElseIf($HTML)
	{
		$rowdata = @()
	}

	$CurrentcpropIndex=0
	ForEach($allclientproperties in $clientprop) 
	{
		for ($v=0;$v -lt $count; $v++)
		{
			#Write-Verbose "$(Get-Date): `t`t`t Processing Enrollment Mode $($Service.DisplayName)";

			If($MSWord -or $PDF)
			{
				## Add the required key/values to the hashtable
				$WordTableRowHash = @{ displayname = $allclientproperties.allclientproperties[$CurrentcpropIndex].displayname;
									   description = $allclientproperties.allclientproperties[$CurrentcpropIndex].description;
									   key = $allclientproperties.allclientproperties[$CurrentcpropIndex].key;
									   value = $allclientproperties.allclientproperties[$CurrentcpropIndex].value;
									   predefined = $allclientproperties.allclientproperties[$CurrentcpropIndex].predefined;
									 }
				## Add the hash to the array
				$cpropWordTable += $WordTableRowHash;
				$CurrentcpropIndex++
			}
			ElseIf($Text)
			{
				Line 0 "Display Name`t: " $allclientproperties.allclientproperties[$CurrentcpropIndex].displayname
				Line 0 "Description`t: " $allclientproperties.allclientproperties[$CurrentcpropIndex].description
				Line 0 "Key`t`t: " $allclientproperties.allclientproperties[$CurrentcpropIndex].key
				Line 0 "Value`t`t: " $allclientproperties.allclientproperties[$CurrentcpropIndex].value
				Line 0 "Predefined`t: " $allclientproperties.allclientproperties[$CurrentcpropIndex].predefined
				Line 0 ""
				$CurrentcpropIndex++
			}
			ElseIf($HTML) 
			{
				$HighlightedCells = $htmlwhite
				$rowdata += @(,(
				$allclientproperties.allclientproperties[$CurrentcpropIndex].displayname,$htmlwhite,
				$allclientproperties.allclientproperties[$CurrentcpropIndex].description,$htmlwhite,
				$allclientproperties.allclientproperties[$CurrentcpropIndex].key,$htmlwhite,
				$allclientproperties.allclientproperties[$CurrentcpropIndex].value,$htmlwhite,
				$allclientproperties.allclientproperties[$CurrentcpropIndex].predefined,$htmlwhite))
				$CurrentcpropIndex++
			}	
		}
	}

	If($MSWord -or $PDF)
	{
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		$Table = AddWordTable -Hashtable $cpropWordTable `
		-Columns displayname, description, key, value, predefined `
		-Headers "Display Name", "Description", "Key", "Value", "Predefined" `
		-Format -155 `
		
		## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
		## AP - Set the size for this table to 10
		SetWordCellFormat -Collection $Table -Size 10
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Color $wdColorWhite -Bold -BackgroundColor $wdColorRoyalBlue;
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 150;
		$Table.Columns.Item(3).Width = 150;
		$Table.Columns.Item(4).Width = 100;
		$Table.Columns.Item(5).Width = 90;
		## IB - Set the required highlighted cells
		SetWordCellFormat -Coordinates $HighlightedCells -Table $Table -Bold -BackgroundColor $wdColorRed -Solid;

		#indent the entire table 1 tab stop
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)
		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($Text)
	{
	}
	ElseIf($HTML)
	{
		$columnHeaders = @(
		'Display Name',($htmlsilver -bor $htmlbold),
		'Description',($htmlsilver -bor $htmlbold),
		'Key',($htmlsilver -bor $htmlbold),
		'Value',($htmlsilver -bor $htmlbold),
		'Predefined',($htmlsilver -bor $htmlbold))
		$msg = "$count Client Properties configuration found"
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 ""
	}
}
#endregion

#region Server Properties
Function ProcessSP
{
	#Bypass certificate verification to enable access with XMS IP Address
	[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

	# Connect to XMS server
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/authentication/login -Body $cred -Headers $headers -Method POST -Verbose:$false
	$headers.add("auth_token",$json.auth_token)

	#List Server Properties configuration
	$srvprop=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/serverproperties" -Headers $headers -Method Get -Verbose:$false

	#Declare an array to collect our result objects
	$resultsarray =@()

	#$count will be the ‘loop counter’
	$count=$srvprop.allewproperties.length
	$chapter++
	Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters Server Properties"

	Try
	{
		$Script:sprop = $srvprop
	}

	Catch
	{
		$Script:sprop = $Null
	}

	If(!$? -or $Null -eq $Script:sprop)
	{
		Write-Warning "No Server Properties were retrieved."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "Warning: No Server Properties were retrieved" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 0 "Warning: No Server Properties were retrieved"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Warning: No Server Properties were retrieved" "" $Null 0 $htmlbold
		}
		Return $False
	}
	ElseIf($? -and $Null -eq $Script:sprop)
	{
		Write-Warning "Server Properties retrieval was successful but no Server Properties were returned."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "Server Properties retrieval was successful but no Server Properties were returned." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 0 "Enrollment Server Properties was successful but no Server Properties were returned."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Server Properties retrieval was successful but no Server Properties were returned." "" $Null 0 $htmlbold
		}
		Return $False
	}
	Else
	{
		If($srvprop -is [array])
		{
			[int]$Script:Numsprop = $count
		}
		Else
		{
			[int]$Script:Numsprop = 1
		}
		Write-Verbose "$(Get-Date): `t`t$($count) Server Properties configuration found"
		Return $True
	}
}

Function ProcessSprop
{
	If($MSWORD -or $PDF)
	{
		$Script:Selection.InsertNewPage()
		Write-Verbose "$(Get-Date): `t`t`tProcessing Server Properties information"
		WriteWordLine 1 0 "Server Properties"
	}
	ElseIf($Text)
	{
		Write-Verbose "$(Get-Date): `t`t`tProcessing Server Properties Mode information"
		Line 0 ""
		Line 0 "Server Properties"
	}
	ElseIf($HTML)
	{
		Write-Verbose "$(Get-Date): `t`t`tProcessing Server Properties Mode information"
		WriteHTMLLine 0 0 ""
		WriteHTMLLine 1 0 "Server Properties"
		WriteHTMLLine 0 0 ""
	}
	
	OutputSprop $Script:sprop $Script:Numsprop
	If($MSWORD -or $PDF)
	{
		OutputSpropNoInternalGridLines $Script:sprop $Script:Numsprop
	}
}

Function OutputSprop
{
	#Bypass certificate verification to enable access with XMS IP Address 
	[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

	#Connect to XMS server
	$headers=@{"Content-Type" = "application/json"}
	$json=Invoke-RestMethod -Uri https://${XMS}:4443/xenmobile/api/v1/authentication/login -Body $cred -Headers $headers -Method POST -Verbose:$false
	$headers.add("auth_token",$json.auth_token)

	#List Server Properties Configuration
	$srvprop=Invoke-RestMethod -Uri "https://${XMS}:4443/xenmobile/api/v1/serverproperties" -Headers $headers -Method Get -Verbose:$false

	# Declare an array to collect our result objects
	$resultsarray =@()

	#$count will be the ‘loop counter’
	$count = $srvprop.allewproperties.length

	Param([object]$srvprop, [int]$count)
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 ""
		WriteWordLine 0 1 "$count Server Properties found on the XMS Server $XMS"
		WriteWordLine 0 1 ""

		## IB - replacement Services table generation utilising AddWordTable function

		## Create an array of hashtables to store our services
		[System.Collections.Hashtable[]] $spropWordTable = @();
		## Create an array of hashtables to store references of cells that we wish to highlight after the table has been added
		[System.Collections.Hashtable[]] $HighlightedCells = @();
		[int] $CurrentspropIndex = 0;
	}
	
	ElseIf($Text)
	{
		Line 0 ""
		Line 0 "`t$count Server Properties configuration found on the XMS Server $XMS"
		Line 0 ""
	}
	
	ElseIf($HTML)
	{
		$rowdata = @()
	}

	$CurrentspropIndex=0
	ForEach($allewproperties in $srvprop)
	{
		for ($v=0;$v -lt $count; $v++)
		{
			If($MSWord -or $PDF)
			{
				## Add the required key/values to the hashtable
				$WordTableRowHash = @{ id = $allewproperties.allewproperties[$CurrentspropIndex].id;
									   name = $allewproperties.allewproperties[$CurrentspropIndex].name;
									   value = $allewproperties.allewproperties[$CurrentspropIndex].value;
									   displayname = $allewproperties.allewproperties[$CurrentspropIndex].displayname;
									   description = $allewproperties.allewproperties[$CurrentspropIndex].description;
									   defaultvalue = $allewproperties.allewproperties[$CurrentspropIndex].defaultvalue;
									 }
				## Add the hash to the array
				$spropWordTable += $WordTableRowHash;
				$CurrentspropIndex++
			}
			ElseIf($Text)
			{
				Line 0 "ID`t`t: " $allewproperties.allewproperties[$CurrentspropIndex].id
				Line 0 "Name`t`t: " $allewproperties.allewproperties[$CurrentspropIndex].name
				Line 0 "Value`t`t: " $allewproperties.allewproperties[$CurrentspropIndex].value
				Line 0 "Display Name`t: " $allewproperties.allewproperties[$CurrentspropIndex].displayname
				Line 0 "Description`t: " $allewproperties.allewproperties[$CurrentspropIndex].description
				Line 0 "Default Value`t: " $allewproperties.allewproperties[$CurrentspropIndex].defaultvalue
				Line 0 ""
				$CurrentspropIndex++
			}
			ElseIf($HTML) 
			{
				$HighlightedCells = $htmlwhite
				$rowdata += @(,(
				$allewproperties.allewproperties[$CurrentspropIndex].id,$htmlwhite,
				$allewproperties.allewproperties[$CurrentspropIndex].name,$htmlwhite,
				$allewproperties.allewproperties[$CurrentspropIndex].value,$htmlwhite,
				$allewproperties.allewproperties[$CurrentspropIndex].displayname,$htmlwhite,
				$allewproperties.allewproperties[$CurrentspropIndex].description,$htmlwhite,
				$allewproperties.allewproperties[$CurrentspropIndex].defaultvalue,$htmlwhite))
				$CurrentspropIndex++
			}	
		}
	}	

	If($MSWord -or $PDF)
	{
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		$Table = AddWordTable -Hashtable $spropWordTable `
		-Columns id, name, value, displayname, description, defaultvalue `
		-Headers "ID","Name", "Value", "Display Name", "Description", "Default Value" `
		-Format -155 `
		
		## AP - Set the size for this table to 10
		SetWordCellFormat -Collection $Table -Size 10
		## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Color $wdColorWhite -Bold -BackgroundColor $wdColorRoyalBlue;
		# Set Columns Width
		$Table.Columns.Item(1).Width = 30;
		$Table.Columns.Item(2).Width = 150;
		$Table.Columns.Item(3).Width = 60;
		$Table.Columns.Item(4).Width = 170;
		$Table.Columns.Item(5).Width = 170;
		$Table.Columns.Item(6).Width = 60;
		
		## IB - Set the required highlighted cells
		SetWordCellFormat -Coordinates $HighlightedCells -Table $Table -Bold -BackgroundColor $wdColorRed -Solid;

		#indent the entire table 1 tab stop
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)
		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($Text)
	{
	}
	ElseIf($HTML)
	{
		$columnHeaders = @(
		'ID',($htmlsilver -bor $htmlbold),
		'Name',($htmlsilver -bor $htmlbold),
		'Value',($htmlsilver -bor $htmlbold),
		'Display Name',($htmlsilver -bor $htmlbold),
		'Description',($htmlsilver -bor $htmlbold),
		'Default Value',($htmlsilver -bor $htmlbold))
		$msg = "$count Server Properties configuration found"
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 ""
	}
}
#endregion
#endregion

#region script setup function
Function ProcessScriptSetup
{
	$script:startTime = Get-Date

	$ComputerName = TestComputerName $ComputerName
}
#endregion

#region script end function
Function ProcessScriptEnd
{
	Write-Verbose "$(Get-Date): Script has completed"
	Write-Verbose "$(Get-Date): "

	#http://poshtips.com/measuring-elapsed-time-in-powershell/
	Write-Verbose "$(Get-Date): Script started: $($Script:StartTime)"
	Write-Verbose "$(Get-Date): Script ended: $(Get-Date)"
	$runtime = $(Get-Date) - $Script:StartTime
	$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds", `
		$runtime.Days, `
		$runtime.Hours, `
		$runtime.Minutes, `
		$runtime.Seconds,
		$runtime.Milliseconds)
	Write-Verbose "$(Get-Date): Elapsed time: $($Str)"

	If($Dev)
	{
		If($SmtpServer -eq "")
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
		}
		Else
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
		}
	}

	If($ScriptInfo)
	{
		$SIFile = "$($pwd.Path)\XMS-TemplateInfo_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
		Out-File -FilePath $SIFile -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Add DateTime    : $($AddDateTime)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Company Name    : $($Script:CoName)" 4>$Null		
		Out-File -FilePath $SIFile -Append -InputObject "Company Address : $($CompanyAddress)" 4>$Null		
		Out-File -FilePath $SIFile -Append -InputObject "Company Email   : $($CompanyEmail)" 4>$Null		
		Out-File -FilePath $SIFile -Append -InputObject "Company Fax     : $($CompanyFax)" 4>$Null		
		Out-File -FilePath $SIFile -Append -InputObject "Company Phone   : $($CompanyPhone)" 4>$Null		
		Out-File -FilePath $SIFile -Append -InputObject "Cover Page      : $($CoverPage)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Dev             : $($Dev)" 4>$Null
		If($Dev)
		{
			Out-File -FilePath $SIFile -Append -InputObject "DevErrorFile    : $($Script:DevErrorFile)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Filename1       : $($Script:FileName1)" 4>$Null
		If($PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Filename2       : $($Script:FileName2)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Folder          : $($Folder)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "From            : $($From)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As HTML    : $($HTML)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As PDF     : $($PDF)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As TEXT    : $($TEXT)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As WORD    : $($MSWORD)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script Info     : $($ScriptInfo)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Port       : $($SmtpPort)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Server     : $($SmtpServer)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Title           : $($Script:Title)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "To              : $($To)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Use SSL         : $($UseSSL)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "User Name       : $($UserName)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "OS Detected     : $($Script:RunningOS)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PoSH version    : $($Host.Version)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSCulture       : $($PSCulture)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSUICulture     : $($PSUICulture)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Word language   : $($Script:WordLanguageValue)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Word version    : $($Script:WordProduct)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script start    : $($Script:StartTime)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Elapsed time    : $($Str)" 4>$Null
	}

	$runtime = $Null
	$Str = $Null
	$ErrorActionPreference = $SaveEAPreference
}
#endregion

#region email function
Function SendEmail
{
	Param([string]$Attachments)
	Write-Verbose "$(Get-Date): Prepare to email"
	
	$emailAttachment = $Attachments
	$emailSubject = $Script:Title
	$emailBody = @"
Hello, <br />
<br />
$Script:Title is attached.
"@ 

	If($Dev)
	{
		Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
	}

	$error.Clear()

	If($UseSSL)
	{
		Write-Verbose "$(Get-Date): Trying to send email using current user's credentials with SSL"
		Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
		-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
		-UseSSL *>$Null
	}
	Else
	{
		Write-Verbose  "$(Get-Date): Trying to send email using current user's credentials without SSL"
		Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
		-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To *>$Null
	}

	$e = $error[0]

	If($e.Exception.ToString().Contains("5.7.57"))
	{
		#The server response was: 5.7.57 SMTP; Client was not authenticated to send anonymous mail during MAIL FROM
		Write-Verbose "$(Get-Date): Current user's credentials failed. Ask for usable credentials."

		If($Dev)
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
		}

		$error.Clear()

		$emailCredentials = Get-Credential -Message "Enter the email account and password to send email"

		If($UseSSL)
		{
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-UseSSL -credential $emailCredentials *>$Null 
		}
		Else
		{
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-credential $emailCredentials *>$Null 
		}

		$e = $error[0]

		If($? -and $Null -eq $e)
		{
			Write-Verbose "$(Get-Date): Email successfully sent using new credentials"
		}
		Else
		{
			Write-Verbose "$(Get-Date): Email was not sent:"
			Write-Warning "$(Get-Date): Exception: $e.Exception" 
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): Email was not sent:"
		Write-Warning "$(Get-Date): Exception: $e.Exception" 
	}
}
#endregion

#region script core
#Script begins
ConnectXMS
ProcessScriptSetup
#Set report Title
[string]$Script:Title = "XenMobile Configuration Report"

#The function SetFileName1andFileName2 needs your script output filename
SetFileName1andFileName2 "XMS-Report"

Write-Verbose "$(Get-Date): Start writing report data"
ProcessSampleText

If(ProcessCert)
{
	ProcessCertificates
	#set orientation
	$Script:Selection.PageSetup.Orientation = $wdOrientLandscape
}
Else
{
	Write-Verbose "$(Get-Date): Certificates cannot be done."
}
$chapter++

If(ProcessLic)
{
	ProcessLicenses
}
Else
{
	Write-Verbose "$(Get-Date): Licenses cannot be done."
}
$chapter++

If(ProcessApp)
{
	ProcessApplications
}
Else
{
	Write-Verbose "$(Get-Date): Applications cannot be done."
}
$chapter++

If(ProcessMDXApp)
{
	ProcessMDXApplications
}
Else
{
	Write-Verbose "$(Get-Date): MDX Applications Settings cannot be done."
}
$chapter++

If(ProcessNS)
{
	ProcessNetScaler
}
Else
{
	Write-Verbose "$(Get-Date): NetScaler cannot be done."
}
$chapter++

If(ProcessAD)
{
	ProcessLDAP
}
Else
{
	Write-Verbose "$(Get-Date): LDAP cannot be done."
}
$chapter++

If(ProcessDev)
{
	ProcessDevices
}
Else
{
	Write-Verbose "$(Get-Date): Devices cannot be done."
}
$chapter++

If(ProcessNSrv)
{
	ProcessNotSrv
}
Else
{
	Write-Verbose "$(Get-Date): Notification Server cannot be done."
}
$chapter++

If(ProcessUG)
{
	ProcessUsersG
}
Else
{
	Write-Verbose "$(Get-Date): Users Groups cannot be done."
}
$chapter++

If(ProcessDgroups)
{
	ProcessDelGroups
}
Else
{
	Write-Verbose "$(Get-Date): Delivery Groups cannot be done."
}
$chapter++

If(ProcessEMode)
{
	ProcessEnrollMode
}
Else
{
	Write-Verbose "$(Get-Date): Enrollment Mode cannot be done."
}
$chapter++

If(ProcessRBAC)
{
	ProcessRBAControl
}
Else
{
	Write-Verbose "$(Get-Date): RBAC cannot be done."
}
$chapter++

If(ProcessCP)
{
	ProcessCprop
}
Else
{
	Write-Verbose "$(Get-Date): RBAC cannot be done."
}
$chapter++

If(ProcessSP)
{
	ProcessSprop
}
Else
{
	Write-Verbose "$(Get-Date): Example output of tables using RBAC cannot be done."
}
#endregion

#region finish script
Write-Verbose "$(Get-Date): Finishing up document"
#end of document processing

###Change the two lines below for your script###
$AbstractTitle = "XenMobile Report"
$SubjectTitle = "XenMobile Report"

UpdateDocumentProperties $AbstractTitle $SubjectTitle

ProcessDocumentOutput

ProcessScriptEnd
#endregion