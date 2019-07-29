# Connect to Exchange Online
Connect-CloudMFA -Tenant Contoso -ExchangeOnline

# Connect to Microsoft Online and AzureAD
Connect-CloudMFA -Tenant Contoso -MSOnline -AzureAD

# With only Tenant parameter, will default to Microsoft Online and AzureAD
Connect-CloudMFA -Tenant Contoso

# Connect to Security and Compliance
Connect-CloudMFA -Tenant Contoso -Compliance

# Connect to SharePoint
Connect-CloudMFA -Tenant Contoso -SharePoint

# For Mailbox Move cmdlets simply type
Connect-CloudMFA Contoso
