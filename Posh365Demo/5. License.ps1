# Connect (if you haven't already)
Connect-CloudMFA mkevin
Connect-CloudMFA mkevin -AzureAD # minimum for licensing cmdlets

# Set Licenses
Set-MailboxMoveLicense -SharePointURL 'https://CoreBTStest.sharepoint.com/sites/mkevin' -ExcelFile 'Batches.xlsx'

# To use CSV instead
Set-MailboxMoveLicense -MailboxCSV 'C:\scripts\Batches.csv'

# Get License Report for user(s)
Get-MailboxMoveLicense -SharePointURL 'https://CoreBTStest.sharepoint.com/sites/mkevin' -ExcelFile 'Batches.xlsx'

# To use CSV instead
Get-MailboxMoveLicense -MailboxCSV 'C:\scripts\Batches.csv' # to use a CSV instead of Excel

# Tenant License Count
Get-MailboxMoveLicenseCount
