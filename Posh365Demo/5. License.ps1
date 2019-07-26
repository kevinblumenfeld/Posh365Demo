# Set Licenses
Set-MailboxMoveLicense -SharePointURL 'https://CoreBTStest.sharepoint.com/sites/mkevin' -ExcelFile 'Batches.xlsx' -Tenant mkevin

# License Reports
Get-MailboxMoveLicense

# Tenant License Count
Get-MailboxMoveLicenseCount
