# Connect (if you haven't already)
Connect-CloudMFA mkevin
Connect-CloudMFA mkevin -AzureAD # minimum for licensing cmdlets

# Set Licenses
Set-MailboxMoveLicense -SharePointURL 'https://CoreBTStest.sharepoint.com/sites/mkevin' -ExcelFile 'Batches.xlsx'
Set-MailboxMoveLicense -MailboxCSV 'C:\scripts\Batches2.csv' # to use a CSV instead of Excel

# Get License Report for user(s)
Get-MailboxMoveLicense -SharePointURL 'https://CoreBTStest.sharepoint.com/sites/mkevin' -ExcelFile 'Batches.xlsx'
Get-MailboxMoveLicense -MailboxCSV 'C:\scripts\Batches2.csv' # to use a CSV instead of Excel

# Tenant License Count
Get-MailboxMoveLicenseCount

$ExcelSplat = @{
    Path                    = 'c:\scripts\LicExcelTest.xlsx'
    FreezeTopRowFirstColumn = $true
    AutoSize                = $true
    ClearSheet              = $true
}
Import-Csv C:\Scripts\356_Licenses.csv | Export-Excel @ExcelSplat -ConditionalText $(
    New-ConditionalText DisplayName White DarkBlue
    New-ConditionalText UserPrincipalName White DarkBlue
    New-ConditionalText AccountSku White DarkBlue
)
