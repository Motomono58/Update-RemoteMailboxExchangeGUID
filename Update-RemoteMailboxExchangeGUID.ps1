
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
        # you can pass the credential using variable ($exoCredential = Get-Credential)
        # then use parameter like so: -credential $exoCredential
        # OR created an encrypted XML (Get-Credential | export-clixml <file.xml>)
        # then use parameter like so: -credential (import-clixml <file.xml>)
        [Parameter(Mandatory=$true)]
        [pscredential]$exoCredential,

        #path to the output directory (eg. c:\scripts\output)
        [Parameter(Mandatory=$true)]
		[string]$outputDirectory,

        #path to the log directory (eg. c:\scripts\logs)
        [Parameter()]
        [string]$logDirectory,
        
        #Sender Email Address
        [Parameter()]
        [string]$sender,

        #Recipient Email Addresses - separate with comma
        [Parameter()]
        [string[]]$recipients,

        #smtpServer
        [Parameter()]
        [string]$smtpServer,

        #smtpPort
        [Parameter()]
        [string]$smtpPort,

        #credential for SMTP server (if applicable)
        [Parameter()]
        [pscredential]$smtpCredential,

        #switch to indicate if SSL will be used for SMTP relay
        [Parameter()]
        [switch]$smtpSSL,

        #Switch to enable email report
        [Parameter()]
        [switch]$sendEmail,

        #Delete older files (in days)
        [Parameter()]
        [int]$removeOldFiles,

        #Domain Controller for use with Exchange Onprem
        [Parameter()]
        [string]$domainController,

        #Exchange ONprem server
        [Parameter(Mandatory=$true)]
        [string]$exchangeServer,

        #Test Mode (if $true, will only simulate but no changes will be applied)
        [Parameter(Mandatory=$true)]
        [boolean]$testMode
)


$script_root = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
#Import Functions
. "$script_root\Functions.ps1"

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

    if (!$smtpServer )
    {
        Write-Host "ERROR: No SMTP Server specified." -ForegroundColor Yellow
        $isAllGood = $false
    }

    if (!$smtpPort )
    {
        Write-Host "ERROR: No SMTP Port specified." -ForegroundColor Yellow
        $isAllGood = $false
    }
}

if ($isAllGood -eq $false)
{
    EXIT
}
#----------------------------------------------------------------------------------------------------

$mailHeader=@'
<!DOCTYPE html>
<html>
<head>
<style>
table {
  font-family: "Century Gothic", sans-serif;
  border-collapse: collapse;
  width: 100%;
}
td, th {
  border: 1px solid #dddddd;
  text-align: left;
  padding: 8px;
}

</style>
</head>
<body>
<table>
'@



#Set Paths-------------------------------------------------------------------------------------------
$today = Get-Date
[string]$fileSuffix = '{0:dd-MMM-yyyy_hh-mm_tt}' -f $today
$logFile = "$($logDirectory)\Log_$($fileSuffix).txt"
$outputHTML = "$($outputDirectory)\Report_$($fileSuffix).html"
$outputCSV = "$($outputDirectory)\Report_$($fileSuffix).csv"

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

#open new Exchange Online Session
Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ': Login to Exchange Online... ' -ForegroundColor Yellow

#Connect to O365 Shell
try 
{
    New-EXOSession $exoCredential
}
catch 
{
    Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": There was an error connecting to Exchange Online. Terminating Script" -ForegroundColor YELLOW
    Stop-TxnLogging
    EXIT
}

#Connect to OnPrem Shell
try 
{
    New-EXSession $exchangeServer
}
catch 
{
    Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": There was an error connecting to Exchange OnPrem. Terminating Script" -ForegroundColor YELLOW
    Stop-TxnLogging
    EXIT
}

$tenantName = (Get-ExoOrganizationConfig).DisplayName

#Retrieve Remote Mailbox Objects with ExchangeGUID 00000000-0000-0000-0000-000000000000
Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Retrieving RemoteMailbox without ExchangeGUID" -ForegroundColor YELLOW
$remoteMailboxes = Get-ExRemoteMailbox -ResultSize Unlimited -Filter {ExchangeGUID -eq "00000000-0000-0000-0000-000000000000"} | Sort-Object Name

if ($remoteMailboxes.Count -gt 0)
{
    Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Comparing with Exchange Online Mailbox" -ForegroundColor YELLOW
	$finalResult = @()
	foreach ($remoteMailbox in $remoteMailboxes)
	{
        $temp = "" | Select-Object Name,UserPrincipalName,OnPremExchangeGUID,OnPremArchiveGUID,OnLineExchangeGUID,OnLineArchiveGUID,Note

        #Get the Exchange Online Mailbox
        $exoMailbox = Get-ExoMailbox $remoteMailbox.Alias -errorAction SilentlyContinue
        
        if ($exoMailbox)
		{			
            $temp.OnLineExchangeGUID = $exoMailbox.ExchangeGuid
            $temp.OnLineArchiveGUID = $exoMailbox.ArchiveGUID
            $temp.Note = "O"
        }
        else 
        {            
			$temp.OnLineExchangeGUID = "NoMailbox"
            $temp.OnLineArchiveGUID = "NoMailbox"
            $temp.Note = "X"
        }

        $temp.Name = $remoteMailbox.Name
        $temp.UserPrincipalName = $remoteMailbox.UserPrincipalName
        $temp.OnPremExchangeGUID = $remoteMailbox.ExchangeGuid
        $temp.OnPremArchiveGUID = $remoteMailbox.ArchiveGUID		
		$finalResult += $temp
    }
    $finalResult = $finalResult | Sort-Object Note
	$finalResult | export-csv -NoTypeInformation $outputCSV
}

if (($finalResult | Where-Object {$_.Note -eq "O"}).Count -gt 0)
{
	Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Updating ExchangeGUID" -ForegroundColor Yellow
	
    $mailSubject = "[$($tenantName)] ExchangeGUID Update Report: " + ('{0:dd-MMM-yyyy hh:mm:ss tt}' -f $Today)

    $htmlBody += $mailHeader

    if ($testMode -eq $true) 
    {
        $htmlBody+="<tr><th>----[TEST MODE]----</th></tr>"
        $htmlBody+="<tr><th>The following RemoteMailbox Exchange GUIDs were updated</th></tr>"        
    }
    else 
    {
        #$htmlBody+="<tr><th>----SUMMARY----</th></tr>"
        $htmlBody+="<tr><th>The following RemoteMailbox Exchange GUIDs were updated</th></tr>" 
    }    
    $htmlBody+="<tr><th>Name</th><th>UPN</th><th>ExchangeGUID.</th></tr>"

	foreach ($result in $finalResult)
	{
        if ($result.OnLineExchangeGUID -eq "NoMailbox")
        {
            Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": ==> $($result.UserPrincipalName), Skipped - NO Mailbox" -ForegroundColor Red
        }
        else
        {
            Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": ==> $($result.UserPrincipalName)" -ForegroundColor Green
            #Update ExchangeGUID for OnPrem Objects
            if ($testMode -eq $false)
            {
                Set-ExRemoteMailbox -Identity $result.UserPrincipalName -ExchangeGUID $result.OnLineExchangeGUID
            }
            $htmlBody += "<tr><td>$($result.Name)</td><td>$($result.UserPrincipalName)</td><td>$($result.OnLineExchangeGUID)</td></tr>"
        }        	
		
	}

    $htmlBody += "</table><a href=""$($scriptInfo.ProjectURI)"">$($scriptInfo.Name)</a> version $($scriptInfo.version)</html>"   
	$htmlBody | out-file $outputHTML
	
    if ($sendEmail -eq $true) 
    {
        Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Sending email to" ($recipients -join ",") -ForegroundColor Green
        $mailParams = @{
            From = $sender
            To = $recipients
            Subject = $mailSubject
            Body = $htmlBody
            smtpServer = $smtpServer
            Port = $smtpPort
            useSSL = $smtpSSL
            Attachments = $outputCSV
            BodyAsHtml = $true
        }

        if ($smtpCredential) 
        {
            $mailParams += @{
                credential = $smtpCredential
            }
        }

        Send-MailMessage @mailParams
    }
}
else 
{
    Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": No Mailbox to be updated. Exiting script." -ForegroundColor Green    
}

#Invoke Housekeeping---------------------------------------------------------------------------------
#if ($enableHousekeeping)
if ($removeOldFiles)
{
	Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Deleting files older than $($removeOldFiles) days" -ForegroundColor Yellow
    Invoke-Housekeeping -folderPath $outputDirectory -daysToKeep $removeOldFiles
    
    if ($logDirectory) {Invoke-Housekeeping -folderPath $logDirectory -daysToKeep $removeOldFiles}
}
#-----------------------------------------------------------------------------------------------
Stop-TxnLogging