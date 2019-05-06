#EDIT THESE VALUES

#[MANDATORY]
#Where is your script file (.PS1) located?
$scriptFile = "C:\Scripts\Update-RemoteMailboxExchangeGUID\Update-RemoteMailboxExchangeGUID.ps1"

#[MANDATORY]
#Where do we save the backup?
$outputDirectory = "C:\Scripts\Update-RemoteMailboxExchangeGUID\Output"

#[MANDATORY]
#Which XML file contains your Office 365 Login? If you don't have this yet, run this: Get-Credential | Export-CliXML <file.xml>
$exoCredentialFile = "C:\Scripts\Update-RemoteMailboxExchangeGUID\credential.xml"

#[MANDATORY]
#your local / onpremise Exchange Server FQDN (eg. exchange1.domain.com)
$exchangeServer = "serverFQDN"

#[MANDATORY]
#Test Mode - if $true, the script will execute but will NOT apply any changes
$testMode = $true

#[OPTIONAL]
#your local / onpremise domain controller FQDN (eg. dc1.domain.com)
$domainController = "serverFQDN"

#[OPTIONAL]
#Where do we put the transcript log?
$logDirectory = "C:\Scripts\Update-RemoteMailboxExchangeGUID\Log"

#======================================
#start EMAIL SECTION
#======================================
#[OPTIONAL]
#Do you want to send the email summary? $true or $false
$sendEmail = $true

#[REQUIRED IF $sendEmail = $true]
#If we will send the email summary, what is the sender email address we should use?
#This must be a valid, existing mailbox and address in Office 365
#The account you use for the Credential File must have "Send As" permission on this mailbox
$sender = "sender@domain.com"

#[REQUIRED IF $sendEmail = $true]
#Who are the recipients?
#Multiple recipients can be added (eg. "recipient1@domain.com","recipient2@domain.com")
$recipients = "recipient1@domain.com","recipient2@domain.com"

#[REQUIRED IF $sendEmail = $true]
#your SMTP relay server
$smtpServer = "smtp.office365.com"

#[REQUIRED IF $sendEmail = $true]
#your SMTP relay server port
$smtpPort = "587"

#[OPTIONAL - use only if your SMTP Relay requires authentication]
#Which XML file contains your SMTP relay authentication? - IF APPLICABLE
#If you don't have this yet, run this: Get-Credential | Export-CliXML <file.xml>
#Or if you are using the same account to login to Office 365, just point to the same XML file
$smtpCredentialFile = "C:\Scripts\Update-RemoteMailboxExchangeGUID\credential.xml"

#[OPTIONAL - use only if SMTP Relay requires SSL]
#Indicate whether or not SSL will be used
$smtpSSL = $true
#======================================
#end EMAIL SECTION
#======================================

#[OPTIONAL]
#If you want to delete older backups, define the age in days.
$removeOldFiles = 30

#======================================
#DO NOT TOUCH THE BELOW CODES
#======================================
$exoCredential = Import-Clixml $exoCredentialFile
if ($smtpCredentialFile) {$smtpCredential = Import-Clixml $smtpCredentialFile}

$params = @{
    outputDirectory = $outputDirectory    
    exoCredential = $exoCredential
    sendEmail = $sendEmail
    testMode = $testMode
    smtpSSL = $smtpSSL
    exchangeServer = $exchangeServer
}

if ($domainController) {$params += @{domainController = $domainController}}
if ($removeOldFiles) {$params += @{removeOldFiles = $removeOldFiles}}
if ($logDirectoryerver) {$params += @{logDirectory = $logDirectory}}
if ($smtpServer) {$params += @{smtpServer = $smtpServer}}
if ($sender) {$params += @{sender = $sender}}
if ($recipients) {$params += @{recipients = $recipients}}
if ($smtpPort) {$params += @{smtpPort = $smtpPort}}
if ($smtpCredentialFile)  {$params += += @{smtpCredential = $smtpCredential}}
#======================================

& "$scriptFile" @params