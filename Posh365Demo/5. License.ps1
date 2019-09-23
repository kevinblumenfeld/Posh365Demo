# Connect (if you haven't already)
Connect-CloudMFA mkevin
Connect-CloudMFA mkevin -AzureAD # minimum for licensing cmdlets

# Set Licenses
Set-MailboxMoveLicense -SharePointURL 'https://CoreBTStest.sharepoint.com/sites/mkevin' -ExcelFile Batches.xlsx

# To use CSV instead
Set-MailboxMoveLicense -MailboxCSV 'C:\scripts\Batches.csv'

# Get License Report for user(s)
Get-MailboxMoveLicense -SharePointURL 'https://CoreBTStest.sharepoint.com/sites/mkevin' -ExcelFile Batches.xlsx

# To use CSV instead
Get-MailboxMoveLicense -MailboxCSV 'C:\scripts\Batches.csv'

# Tenant License Count
Get-MailboxMoveLicenseCount

# Detailed Tenant License Report. Per User, Per Sku, Per Option.
Connect-CloudMFA -Tenant Contoso -MSOnline
Get-MailboxMoveLicenseReport -Path 'C:\scripts'
