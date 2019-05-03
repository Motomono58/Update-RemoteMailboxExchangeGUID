
<#PSScriptInfo

.VERSION 1.0

.GUID 544340d7-051b-4479-ac66-7b8eea3ca7d2

.AUTHOR june

.COMPANYNAME

.COPYRIGHT

.TAGS

.LICENSEURI

.PROJECTURI

.ICONURI

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES


.PRIVATEDATA

#> 



<# 

.DESCRIPTION 
Update GUID

#> 
Param(
        # office 365 credential
        # you can pass the credential using variable ($credential = Get-Credential)
        # then use parameter like so: -credential $credential
        # OR created an encrypted XML (Get-Credential | export-clixml <file.xml>)
        # then use parameter like so: -credential (import-clixml <file.xml>)
        [Parameter(Mandatory=$true,Position=0)]
        [pscredential]$credential,

        #path to the output directory (eg. c:\scripts\output)
        [Parameter(Mandatory=$true,Position=1)]
		[string]$outputDirectory,
		
		#limit the result
        [Parameter(Mandatory=$true,position=2)]
		$resultSizeLimit,

        #path to the log directory (eg. c:\scripts\logs)
        [Parameter()]
        [string]$logDirectory,
        
        #Sender Email Address
        [Parameter()]
        [string]$sender,

        #Recipient Email Addresses - separate with comma
        [Parameter()]
        [string[]]$recipients,

        #Switch to enable email report
        [Parameter()]
        [switch]$sendEmail,

        #Delete older files (in days)
        [Parameter()]
		[int]$removeOldFiles 
)


$script_root = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition

Stop-TxnLogging
Clear-Host
$scriptInfo = Test-ScriptFileInfo -Path $MyInvocation.MyCommand.Definition


#parameter check ----------------------------------------------------------------------------------------------------
$isAllGood = $true

if ($sendEmail)
{
    if (!$sender)
    {
        Write-Host "ERROR: A valid sender email address is not specified." -ForegroundColor Yellow
        $isAllGood = $false
    }

    if (!$recipients)
    {
        Write-Host "ERROR: No recipients specified." -ForegroundColor Yellow
        $isAllGood = $false
    }
}

if ($isAllGood -eq $false)
{
    EXIT
}
#----------------------------------------------------------------------------------------------------

#Import Functions
. "$script_root\Functions.ps1"

#Set Paths-------------------------------------------------------------------------------------------
$Today=Get-Date
[string]$fileSuffix = '{0:dd-MMM-yyyy_hh-mm_tt}' -f $Today
$logFile = "$($logDirectory)\Log_$($fileSuffix).txt"

#Create folders if not found
if ($logDirectory)
{
    if (!(Test-Path $logDirectory)) 
    {
        New-Item -ItemType Directory -Path $logDirectory | Out-Null
        #start transcribing----------------------------------------------------------------------------------
        Start-TxnLogging $logFile
        #----------------------------------------------------------------------------------------------------
    }
	else
	{
		Start-TxnLogging $logFile
	}
}

if (!(Test-Path $outputDirectory)) 
{
	New-Item -ItemType Directory -Path $outputDirectory | Out-Null
}
#----------------------------------------------------------------------------------------------------

#<mail variables
$subject = "ExchangeGUID Update Report"
$smtpServer = "smtp.office365.com"
$smtpPort = "587"
#mail variables>