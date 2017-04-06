# Get-MailboxPermissionReport.ps1
Dump mailbox folder permissions to CSV file

## Description
This script exports all mailbox folder permissions for mailboxes of type "UserMailbox".
    
The permissions are exported to a local CSV file

The script is intended to run from within an active Exchange 2013 Management Shell session.
## Parameters
### CsvFileName

## Outputs
CSV report containing mailbox folder permissions.

## Examples
```
.\Get-MailboxPermissionsReport-ps1 -CsvFileName export.csv
```
Export mailbox permissions to export.csv


## TechNet Gallery
Find the script at TechNet Gallery
* TBD

## Credits
Written by: Thomas Stensitzki

## Social 

* My Blog: http://justcantgetenough.granikos.eu
* Archived Blog: http://www.sf-tools.net/
* Twitter:	https://twitter.com/stensitzki
* LinkedIn:	http://de.linkedin.com/in/thomasstensitzki
* Github:	https://github.com/Apoc70

For more Office 365, Cloud Security and Exchange Server stuff checkout services provided by Granikos

* Blog:     http://blog.granikos.eu/
* Website:	https://www.granikos.eu/en/
* Twitter:	https://twitter.com/granikos_de

Additional Credits:
* This script is based on Mr Tony Redmonds blog post http://thoughtsofanidlemind.com/2014/09/05/reporting-delegate-access-to-exchange-mailboxes/ 
