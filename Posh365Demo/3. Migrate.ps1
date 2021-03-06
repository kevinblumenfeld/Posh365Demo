# Connect
Connect-CloudMFA mkevin

# New
New-MailboxMove -SharePointURL 'https://CoreBTStest.sharepoint.com/sites/mkevin' -ExcelFile 'Batches.xlsx' -Tenant mkevin -RemoteHost mail.oktakevin.com
New-MailboxMove -MailboxCSV 'c:\scripts\Batches.csv' -RemoteHost mail.oktakevin.com -Tenant mkevin # To use CSV

# Get
Get-MailboxMove
Get-MailboxMove -IncludeCompleted

# Set
Set-MailboxMove -LargeItemLimit 100 -BadItemLimit 200

# Complete
Complete-MailboxMove
Complete-MailboxMove -Schedule

# Suspend and Resume
Suspend-MailboxMove
Resume-MailboxMove
Resume-MailboxMove -DontAutoComplete

# Remove
Remove-MailboxMove
