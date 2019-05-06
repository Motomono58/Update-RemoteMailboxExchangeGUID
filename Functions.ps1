
<#PSScriptInfo

.VERSION 1.1

.GUID ef8d8a67-099d-4250-b6ea-bccb73d83a08

.AUTHOR June Castillote

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
Added Exchange On-Prem connection function

.PRIVATEDATA

#>

<# 

.DESCRIPTION 
 Helper Functions

#>

#Function to connect to Exchang Online Shell
Function New-EXOSession
{
    [CmdletBinding()]
    param(
        [parameter(mandatory=$true,position=0)]
        [PSCredential] $exoCredential
    )

    Remove-PSSession -Name "ExchangeOnline" -Confirm:$false -ErrorAction SilentlyContinue
    $EXOSession = New-PSSession -Name "ExchangeOnline" -ConfigurationName "Microsoft.Exchange" -ConnectionUri 'https://ps.outlook.com/powershell' -Credential $exoCredential -Authentication Basic -AllowRedirection -WarningAction SilentlyContinue
    Import-PSSession $EXOSession -AllowClobber -DisableNameChecking -Prefix Exo | out-null
}

#Function to connect to Exchange OnPrem Shell
Function New-ExSession()
{
    [CmdletBinding()]
    param(
        [parameter(mandatory=$true,position=0)]
        [string] $exchangeServer
    )
    Remove-PSSession -Name "ExchangeOnPrem" -Confirm:$false -ErrorAction SilentlyContinue
	$EXSession = New-PSSession -Name "ExchangeOnPrem" -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$($exchangeServer)/PowerShell/" -Authentication Kerberos
	Import-PSSession $EXSession -AllowClobber -DisableNameChecking -Prefix Ex | out-null
}

#Function to compress file (ps 4.0)
Function New-ZipFile
{
	[CmdletBinding()] 
    param ( 
        [Parameter(Mandatory=$true,position=0)] 
        [string]$fileToZip,    
        
        [Parameter(Mandatory=$true,position=1)]
        [string]$destinationZip
	)
	Add-Type -assembly System.IO.Compression
	Add-Type -assembly System.IO.Compression.FileSystem
	[System.IO.Compression.ZipArchive]$outZipFile = [System.IO.Compression.ZipFile]::Open($destinationZip, ([System.IO.Compression.ZipArchiveMode]::Create))
	[System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($outZipFile, $fileToZip, (Split-Path $fileToZip -Leaf)) | out-null
	$outZipFile.Dispose()
}

#Function to delete old files based on age
Function Invoke-Housekeeping
{
    [CmdletBinding()] 
    param ( 
        [Parameter(Mandatory=$true,position=0)] 
        [string]$folderPath,
    
		[Parameter(Mandatory=$true,position=1)]
		[int]$daysToKeep
    )
    
    $datetoDelete = (Get-Date).AddDays(-$daysToKeep)
    $filesToDelete = Get-ChildItem $FolderPath | Where-Object { $_.LastWriteTime -lt $datetoDelete }

    if (($filesToDelete.Count) -gt 0) {	
		foreach ($file in $filesToDelete) {
            Remove-Item -Path ($file.FullName) -Force -ErrorAction SilentlyContinue
		}
	}	
}

#Function to Stop Transaction Logging
Function Stop-TxnLogging
{
	$txnLog=""
	Do {
		try {
			Stop-Transcript | Out-Null
		} 
		catch [System.InvalidOperationException]{
			$txnLog="stopped"
		}
    } While ($txnLog -ne "stopped")
}

#Function to Start Transaction Logging
Function Start-TxnLogging
{
    param 
    (
        [Parameter(Mandatory=$true,Position=0)]
        [string]$logDirectory
    )
	Stop-TxnLogging
    Start-Transcript $logDirectory -Append
}