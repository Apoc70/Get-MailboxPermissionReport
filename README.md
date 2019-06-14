# Get-MailboxPermissionReport.ps1

Dump mailbox folder permissions to CSV file

## Description

This script exports all mailbox folder permissions for mailboxes of type "UserMailbox".

The permissions are exported to a local CSV file.

The script is intended to run from within an active Exchange 2013 Management Shell session.

## Parameters

### MailboxId

The mailbox id for filtering mailboxes, default *

### CsvFileName

The file name for the export CSV file, default MailboxPermissions.csv

## Examples

``` PowerShell
.\Get-MailboxPermissionsReport-ps1 -MailboxId "John*" -CsvFileName export.csv
```

Export permissions of all mailboxes starting with "John*" to file export.csv

## Note

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## TechNet Gallery

Find the script at TechNet Gallery

* [https://gallery.technet.microsoft.com/Export-all-user-mailbox-155a33de](https://gallery.technet.microsoft.com/Export-all-user-mailbox-155a33de)

## Credits

Written by: Thomas Stensitzki

Stay connected:

* My Blog: [http://justcantgetenough.granikos.eu](http://justcantgetenough.granikos.eu)
* Twitter: [https://twitter.com/stensitzki](https://twitter.com/stensitzki)
* LinkedIn: [http://de.linkedin.com/in/thomasstensitzki](http://de.linkedin.com/in/thomasstensitzki)
* Github: [https://github.com/Apoc70](https://github.com/Apoc70)

For more Office 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

* Blog: [http://blog.granikos.eu](http://blog.granikos.eu)
* Website: [https://www.granikos.eu/en/](https://www.granikos.eu/en/)
* Twitter: [https://twitter.com/granikos_de](https://twitter.com/granikos_de)

Additional Credits:

* This script is based on Mr Tony Redmonds blog post [http://thoughtsofanidlemind.com/2014/09/05/reporting-delegate-access-to-exchange-mailboxes](http://thoughtsofanidlemind.com/2014/09/05/reporting-delegate-access-to-exchange-mailboxes)