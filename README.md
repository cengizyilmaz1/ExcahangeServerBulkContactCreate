# ExcahangeServerBulkContactCreate

Create Mail-Enabled Contacts from CSV
This PowerShell script creates mail-enabled contacts in Microsoft Exchange using data from a CSV file. It sets each contact’s ExternalEmailAddress (taken from the CSV’s mail column) and generates EmailAddresses (proxy address) and MailNickname from a user-defined prefix, domain, and a sequential counter.

Features
CSV-based Creation: Reads contact data from a CSV file (comma-delimited).
Dynamic Proxy Address: Uses a customizable prefix and domain to create sequential proxy addresses (e.g. ext.mail01@contoso.com, ext.mail02@contoso.com).
Auto-Truncation Option: Lets you truncate Display Names longer than 64 characters automatically or skip them.
Configurable OU: You specify the Organizational Unit (OU) where the mail contacts will be created.
Error Handling & Summary: Shows real-time success or failure for each contact and displays a summary at the end.

Requirements
Exchange Management Shell (EMS):
Must be run in the Exchange Management Shell or a PowerShell environment where New-MailContact and Set-MailContact cmdlets are available.

Sufficient Permissions:
The account running the script needs the rights to create and modify mail contacts in the target OU.

CSV File:
Must be comma-delimited (.csv) with the following header columns:
Display Name
First
Last
mail
Setup and Usage
Download or Copy the Script

Save the script (e.g., CreateMailContactsFromCSV.ps1) to a folder of your choice.
Open Exchange Management Shell

Start the Exchange Management Shell (or a suitable PowerShell session with Exchange cmdlets loaded).
Check Execution Policy (If Needed)

If your PowerShell execution policy blocks the script, you can temporarily set it (only for the current session):
Set-ExecutionPolicy RemoteSigned -Scope Process


Navigate to the folder containing the script:
cd C:\Scripts
.\CreateMailContactsFromCSV.ps1
Script Prompts

Target OU: You’ll be asked for the Distinguished Name of the target OU (e.g., OU=Contacts,DC=company,DC=local).
CSV Path: Enter the full path to the CSV file (e.g., C:\Data\contacts.csv).
Auto-Truncate: Choose whether to truncate Display Names if over 64 characters (Y or N).
Proxy Prefix: Enter the prefix for generating proxy addresses (e.g., ext.mail).
Proxy Domain: Enter the domain name (e.g., contoso.com).
The script will then read each row in the CSV and create a mail-enabled contact in Exchange.

CSV Format
Your CSV file must include these headers exactly (comma-delimited):

Display Name
First
Last
mail
Below is an example CSV snippet (two contacts):

"Display Name","First","Last","mail"
"John Smith","John","Smith","[email protected]"
"Jane Doe","Jane","Doe","[email protected]"


The Display Name is used for the contact’s display name in Exchange (and may be truncated if too long).
The mail field is used for the ExternalEmailAddress property.
The proxy address (EmailAddresses) and MailNickname are derived from the user-input prefix/domain plus a sequential counter.
Example

  .\CreateMailContactsFromCSV.ps1
  Enter the target OU's Distinguished Name (e.g., OU=Contacts,DC=fixcloud,DC=com,DC=tr): OU=Contacts,DC=company,DC=local
  Enter the full path to the CSV file (e.g., C:\Data\contacts.csv): C:\Data\contacts.csv
  Auto-truncate Display Names longer than 64 characters? (Y/N): Y
  Enter the proxy address prefix (e.g., ext.mail): ext.mail
  Enter the proxy address domain (e.g., contoso.com): contoso.com
  Process Output:
  Success: Created mail contact 'John Smith'
         ExternalEmailAddress (target) = [email protected]
         EmailAddresses (proxy) = ext.mail01@contoso.com
         MailNickname (alias)   = ext.mail01_contoso.com
----------------------------------------
Operation Summary:
Total contacts processed: 2
Successfully created: 2
Failed to create: 0
Operation completed.
