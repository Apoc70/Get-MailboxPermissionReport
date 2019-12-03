<#
    .SYNOPSIS
    Dump mailbox folder permissions to CSV file
   
    Thomas Stensitzki
	
    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
    Version 1.4, 2019-06-14

    Ideas, comments and suggestions to support@granikos.eu 

    This script is based on Mr Tony Redmonds blog post http://thoughtsofanidlemind.com/2014/09/05/reporting-delegate-access-to-exchange-mailboxes/ 
 
    .LINK  
    http://www.granikos.eu/en/scripts 
	
    .DESCRIPTION
    This script exports all mailbox folder permissions for mailboxes of type "UserMailbox".
    
    The permissions are exported to a local CSV file

    The script is inteded to run from within an active Exchange 2013 Management Shell session.

    .NOTES 
    Requirements 
    - Windows Server 2012 or newer
    - Exchange Server Management Shell (EMS) 2010+

    Revision History 
    -------------------------------------------------------------------------------- 
    1.0 | Initial community release
    1.1 | Minor PowerShell fix
    1.2 | Minor PowerShell changes
    1.3 | MailboxId parameter
    1.4 | Fix to ensure mailbox uniqueness using the alias property
	
    .PARAMETER MailboxId
    Mailbox filter, default *

    .PARAMETER CsvFileName
    CSV file name, default MailboxPermissions.csv

    .EXAMPLE
    Export mailbox permissions to export.csv

    .\Get-MailboxPermissionsReport-ps1 -CsvFileName export.csv

#>
[CmdletBinding()]
Param(
  [string]$MailboxId = '*',
  [string]$CsvFileName = 'MailboxPermissions.csv'
  
)

$ScriptDir = Split-Path -Path $script:MyInvocation.MyCommand.Path
$ScriptName = $MyInvocation.MyCommand.Name

# build CSV full path to store CSV in script directory
$OutputFile = Join-Path -Path $ScriptDir -ChildPath $CsvFileName

Write-Verbose -Message $OutputFile

# Fetch mailboxes of type UserMailbox only
$Mailboxes = Get-Mailbox -RecipientTypeDetails 'UserMailbox' -Identity $MailboxId -ResultSize Unlimited | Sort-Object

$result = @()

# counter for progress bar
$MailboxCount = ($Mailboxes | Measure-Object).Count
$count = 1

ForEach ($Mailbox in $Mailboxes) { 

  $Alias = '' + $Mailbox.Alias # Use Alias property instead of name to ensure 'uniqueness' passed on to Get-MailboxFolderStatistics
  
  $DisplayName = ('{0} ({1})' -f $Mailbox.DisplayName, $Mailbox.Name)

  $activity = ('Working... [{0}/{1}]' -f $count, $mailboxCount)
  $status = ('Getting folders for mailbox: {0}' -f $DisplayName)
  Write-Progress -Status $status -Activity $activity -PercentComplete (($count/$MailboxCount)*100) 
    
  # Fetch folders
  $Folders = @('\')
  $Folders += Get-MailboxFolderStatistics $Alias | Select-Object -Skip 1 | ForEach-Object {$_.folderpath} | ForEach-Object{$_.replace('/','\')}

  ForEach ($Folder in $Folders) {
  
    # build folder key to fetch mailbox folder permissions
    $FolderKey = $Alias + ':' + $Folder
    
    # fetch mailbox folder permissions
    $Permissions = Get-MailboxFolderPermission -Identity $FolderKey -ErrorAction SilentlyContinue
    
    # store results in variable
    $result += $Permissions | Where-Object {$_.User -notlike 'Default' -and $_.User -notlike 'Anonymous' -and $_.AccessRights -notlike 'None' -and $_.AccessRights -notlike 'Owner' } | Select-Object -Property @{name='Mailbox';expression={$DisplayName}}, FolderName, @{name='Identity';expression={$Folder}}, @{name='User';expression={$_.User -join ','}}, @{name='AccessRights';expression={$_.AccessRights -join ','}}
  }
    
  # Increment counter
  $count++
}

# Export to CSV
$result | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8 -Delimiter ';' -Force
