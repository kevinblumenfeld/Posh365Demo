# Connect (if you haven't already)
Connect-CloudMFA mkevin

# Get Mailbox and Folder Permissions for requested users
Get-MailboxMovePermission -SharePointURL 'https://mkevin.sharepoint.com/sites/Kevin' -ExcelFile 'Batches.xlsx' -Tenant mkevin

# Add to Exchange Online Mailbox and Folder Permissions for requested users
Add-MailboxMovePermission -SharePointURL 'https://mkevin.sharepoint.com/sites/Kevin' -ExcelFile 'Batches.xlsx' -Tenant mkevin
