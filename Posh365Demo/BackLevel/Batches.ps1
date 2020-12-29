[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force -ErrorAction SilentlyContinue
function Get-MailboxMoveOnPremisesMailboxReport {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $ReportPath
    )
    end {
        New-Item -ItemType Directory -Path $ReportPath -ErrorAction SilentlyContinue
        $BatchesFile = Join-Path $ReportPath 'Batches.csv'
        $Select = @(
            'BatchName', 'DisplayName', 'Enabled', 'OrganizationalUnit', 'IsMigrated', 'CompleteBatchDate'
            'CompleteBatchTimePT', 'LicenseGroup', 'EnableArchive', 'ConvertToShared', 'MailboxGB'
            'ArchiveGB', 'DeletedGB', 'TotalGB', 'LastLogonTime', 'ItemCount', 'UserPrincipalName'
            'PrimarySmtpAddress', 'AddressBookPolicy', 'RetentionPolicy', 'Alias', 'Database'
            'ServerName', 'OU', 'Department', 'Office', 'SamAccountName', 'RecipientTypeDetails'
            'UMEnabled', 'ForwardingAddress', 'ForwardingRecipientType', 'ForwardingSmtpAddress'
            'DeliverToMailboxAndForward', 'ExchangeGuid'
        )
        Get-MailboxMoveOnPremisesReportHelper | Select-Object $Select | Export-Csv $BatchesFile -NoTypeInformation -Encoding UTF8

        $ExcelSplat = @{
            Path                    = (Join-Path $ReportPath 'Batches.xlsx')
            TableStyle              = 'Medium2'
            FreezeTopRowFirstColumn = $true
            AutoSize                = $true
            BoldTopRow              = $true
            ClearSheet              = $true
            WorksheetName           = 'Batches'
            ErrorAction             = 'stop'
        }
        try {
            $BatchesFile | Where-Object { $_ } | ForEach-Object { Import-Csv $_ | Export-Excel @ExcelSplat }
        }
        catch {
            $_.Exception.Message
        }
    }
}

function Get-MailboxMoveOnPremisesReportHelper {
    [CmdletBinding()]
    param (

    )
    end {
        $RecHash = Get-MailboxMoveRecipientHash
        $ADHash = Get-ADHash
        $MailboxList = Get-Mailbox -ResultSize Unlimited -IgnoreDefaultScope
        $MailboxList = $MailboxList | Where-Object { $_.RecipientTypeDetails -ne 'DiscoveryMailbox' }
        foreach ($Mailbox in $MailboxList) {
            Write-Verbose "Mailbox`t$($Mailbox.DisplayName)"
            $Statistic = $Mailbox | Get-ExchangeMailboxStatistics
            $PSHash = @{
                BatchName                  = 'zNoBatch'
                DisplayName                = $Mailbox.DisplayName
                OrganizationalUnit         = $Mailbox.OrganizationalUnit
                IsMigrated                 = ''
                CompleteBatchDate          = ''
                CompleteBatchTimePT        = ''
                LicenseGroup               = ''
                EnableArchive              = ''
                ConvertToShared            = ''
                MailboxGB                  = $Statistic.MailboxGB
                ArchiveGB                  = $Statistic.ArchiveGB
                DeletedGB                  = $Statistic.DeletedGB
                TotalGB                    = $Statistic.TotalGB
                LastLogonTime              = $Statistic.LastLogonTime
                ItemCount                  = $Statistic.ItemCount
                UserPrincipalName          = $Mailbox.UserPrincipalName
                PrimarySmtpAddress         = $Mailbox.PrimarySmtpAddress
                AddressBookPolicy          = $Mailbox.AddressBookPolicy
                RetentionPolicy            = $Mailbox.RetentionPolicy
                Enabled                    = $ADHash[$Mailbox.UserPrincipalName]['Enabled']
                SamAccountName             = $ADHash[$Mailbox.UserPrincipalName]['SamAccountName']
                Alias                      = $Mailbox.Alias
                Database                   = $Mailbox.Database
                ServerName                 = $Mailbox.ServerName
                OU                         = ($Mailbox.DistinguishedName -replace '^.+?,(?=(OU|CN)=)')
                Department                 = $ADHash[$Mailbox.UserPrincipalName]['Department']
                Office                     = $Mailbox.Office
                RecipientTypeDetails       = $Mailbox.RecipientTypeDetails
                UMEnabled                  = $Mailbox.UMEnabled
                ForwardingSmtpAddress      = $Mailbox.ForwardingSmtpAddress
                DeliverToMailboxAndForward = $Mailbox.DeliverToMailboxAndForward
                ExchangeGuid               = $Mailbox.ExchangeGuid
            }
            if ($Mailbox.ForwardingAddress) {
                $Distinguished = Convert-CanonicalToDistinguished -CanonicalName $Mailbox.ForwardingAddress
                $PSHash['ForwardingAddress'] = $RecHash[$Distinguished].PrimarySmtpAddress
                $PSHash['ForwardingRecipientType'] = $RecHash[$Distinguished].RecipientTypeDetails
            }
            else {
                $PSHash['ForwardingAddress'] = ''
                $PSHash['ForwardingRecipientType'] = ''
            }
            [PSCustomObject]$PSHash
        }
    }
}

function Get-ExchangeMailboxStatistics {
    [CmdletBinding()]
    param (

        [Parameter(ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            Mandatory = $false)]
        $MailboxList
    )
    Begin {

    }
    Process {
        foreach ($Mailbox in $MailboxList) {
            $ArchiveGB = Get-MailboxStatistics -identity ($Mailbox.PrimarySmtpAddress).ToString() -Archive -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | ForEach-Object {
                [Math]::Round([Double]($_.TotalItemSize -replace '^.*\(| .+$|,') / 1GB, 4)
            }
            Get-MailboxStatistics -identity ($Mailbox.PrimarySmtpAddress).ToString() -WarningAction SilentlyContinue | Select-Object @(
                'DisplayName'
                @{
                    Name       = 'PrimarySmtpAddress'
                    Expression = { $Mailbox.PrimarySmtpAddress }
                }
                @{
                    Name       = 'UserPrincipalName'
                    Expression = { $Mailbox.UserPrincipalName }
                }
                @{
                    Name       = 'MailboxGB'
                    Expression = {
                        [Math]::Round([Double]($_.TotalItemSize -replace '^.*\(| .+$|,') / 1GB, 4)
                    }
                }
                @{
                    Name       = 'ArchiveGB'
                    Expression = { $ArchiveGB }
                }
                @{
                    Name       = 'DeletedGB'
                    Expression = {
                        [Math]::Round([Double]($_.TotalDeletedItemSize -replace '^.*\(| .+$|,') / 1GB, 4)
                    }
                }
                @{
                    Name       = 'TotalGB'
                    Expression = {
                        [Math]::Round([Double]($_.TotalItemSize -replace '^.*\(| .+$|,') / 1GB, 4) + $ArchiveGB
                    }
                }
                'LastLogonTime'
                'ItemCount'
            )
        }
    }
    End {

    }
}

function Convert-CanonicalToDistinguished {
    [cmdletbinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $CanonicalName
    )
    end {
        $nameTranslate = New-Object -ComObject NameTranslate
        # $nameTranslate.Init(3,  '')
        # To PS2:
        [__ComObject].InvokeMember('Init', 'InvokeMethod', $null, $nameTranslate, @(3, ''), $null, (Get-Culture), $null)
        # Get an identity using the canonicalName
        # $nameTranslate.Set(2, $canonicalName)
        # To PS2:
        [__ComObject].InvokeMember('Set', 'InvokeMethod', $null, $nameTranslate, @(2, $canonicalName), $null, (Get-Culture), $null)
        # Convert the identity to a DistinguishedName
        # $nameTranslate.Get(1)
        # To PS2:
        [__ComObject].InvokeMember('Get', 'InvokeMethod', $null, $nameTranslate, @(1), $null, (Get-Culture), $null)
    }
}

function Get-MailboxMoveRecipientHash {
    [CmdletBinding()]
    param
    (
    )
    end {
        $RecipientHash = @{ }
        $RecipientList = Get-Recipient -ResultSize Unlimited
        foreach ($Recipient in $RecipientList) {
            $RecipientHash[$Recipient.DistinguishedName] = @{
                PrimarySMTPAddress   = $Recipient.PrimarySMTPAddress
                RecipientTypeDetails = $Recipient.RecipientTypeDetails
            }
        }
        $RecipientHash
    }
}

function Connect-Exchange {
    [CmdletBinding()]
    param
    (
        [Parameter()]
        [string]
        $Server,

        [Parameter()]
        [Switch]
        $DeleteExchangeCreds,

        [Parameter()]
        [Switch]
        $DontViewEntireForest
    )

    $CredFile = Join-Path $Env:USERPROFILE ConnectExchange.xml
    if ($DeleteExchangeCreds) { Remove-Item $CredFile -Force }

    if (-not ($null = Test-Path $CredFile)) {
        [System.Management.Automation.PSCredential]$Credential = Get-Credential -Message 'Enter on-premises Exchange username and password'
        [System.Management.Automation.PSCredential]$Credential | Export-Clixml -Path $CredFile
        [System.Management.Automation.PSCredential]$Credential = Import-Clixml -Path $CredFile
    }
    else {
        [System.Management.Automation.PSCredential]$Credential = Import-Clixml -Path $CredFile
    }
    $SessionSplat = @{
        Name              = "OnPremExchange"
        ConfigurationName = 'Microsoft.Exchange'
        ConnectionUri     = ("http://" + $Server + "/PowerShell/")
        Authentication    = 'Kerberos'
        Credential        = $Credential
    }
    $Session = New-PSSession @SessionSplat
    $SessionModule = Import-PSSession -AllowClobber -DisableNameChecking -Session $Session
    $null = Import-Module $SessionModule -Global -DisableNameChecking -Force
    if (-not $DontViewEntireForest) {
        Set-ADServerSettings -ViewEntireForest:$True
    }
    Write-Host "Connected to Exchange Server: $Server" -ForegroundColor Green
}

function New-ADSIPrincipalContext {
    <#
    .NOTES
        https://github.com/lazywinadmin/ADSIPS

    .LINK
        https://msdn.microsoft.com/en-us/library/system.directoryservices.accountmanagement.principalcontext(v=vs.110).aspx
    #>

    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType('System.DirectoryServices.AccountManagement.PrincipalContext')]
    param
    (
        [Alias("RunAs")]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory = $true)]
        [System.DirectoryServices.AccountManagement.ContextType]$ContextType,

        $DomainName = [System.DirectoryServices.ActiveDirectory.Domain]::Getcurrentdomain(),

        $Container,

        [System.DirectoryServices.AccountManagement.ContextOptions[]]$ContextOptions
    )

    begin {
        $ScriptName = (Get-Variable -name MyInvocation -Scope 0 -ValueOnly).MyCommand
        Write-Verbose -Message "[$ScriptName] Add Type System.DirectoryServices.AccountManagement"
        Add-Type -AssemblyName System.DirectoryServices.AccountManagement
    }
    process {
        try {
            switch ($ContextType) {
                "Domain" {
                    $ArgumentList = $ContextType, $DomainName
                }
                "Machine" {
                    $ArgumentList = $ContextType, $ComputerName
                }
                "ApplicationDirectory" {
                    $ArgumentList = $ContextType
                }
            }

            if ($PSBoundParameters['Container']) {
                $ArgumentList += $Container
            }

            if ($PSBoundParameters['ContextOptions']) {
                $ArgumentList += $($ContextOptions)
            }

            if ($PSBoundParameters['Credential']) {
                # Query the specified domain or current if not entered, with the specified credentials
                $ArgumentList += $($Credential.UserName), $($Credential.GetNetworkCredential().password)
            }

            if ($PSCmdlet.ShouldProcess($DomainName, "Create Principal Context")) {
                # Query
                New-Object -TypeName System.DirectoryServices.AccountManagement.PrincipalContext -ArgumentList $ArgumentList
            }
        } #try
        catch {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    } #process
}

function Get-ADSIUser {
    <#
    .NOTES
        https://github.com/lazywinadmin/ADSIPS
    .LINK
        https://msdn.microsoft.com/en-us/library/System.DirectoryServices.AccountManagement.UserPrincipal(v=vs.110).aspx
    #>

    [CmdletBinding(DefaultParameterSetName = "All")]
    [OutputType('System.DirectoryServices.AccountManagement.UserPrincipal')]
    param
    (
        [Parameter(Mandatory = $true, Position = 0, ParameterSetName = "Identity")]
        [string]$Identity,

        [Alias("RunAs")]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty,

        [String]$DomainName,

        [Parameter(Mandatory = $true, ParameterSetName = "LDAPFilter")]
        [string]$LDAPFilter,

        [Parameter(ParameterSetName = "LDAPFilter")]
        [Parameter(ParameterSetName = "All")]
        [Switch]$NoResultLimit

    )

    begin {
        Add-Type -AssemblyName System.DirectoryServices.AccountManagement

        # Create Context splatting
        $ContextSplatting = @{ ContextType = "Domain" }

        if ($PSBoundParameters['Credential']) {
            $ContextSplatting.Credential = $Credential
        }
        if ($PSBoundParameters['DomainName']) {
            $ContextSplatting.DomainName = $DomainName
        }

        $Context = New-ADSIPrincipalContext @ContextSplatting
    }
    process {
        if ($Identity) {
            Write-Verbose -Message "Identity"
            try {
                [System.DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($Context, $Identity)
            }
            catch {
                if ($_.Exception.Message.ToString().EndsWith('"Multiple principals contain a matching Identity."')) {
                    $errorMessage = "[Get-ADSIUser] On line $($_.InvocationInfo.ScriptLineNumber) - We found multiple entries for Identity: '$($Identity)'. Please specify a samAccountName, or something more specific."
                    $MultipleEntriesFoundException = [System.Exception]::new($errorMessage)
                    throw $MultipleEntriesFoundException
                }
                else {
                    $PSCmdlet.ThrowTerminatingError($_)
                }
            }
        }
        elseif ($PSBoundParameters['LDAPFilter']) {

            # Directory Entry object
            $DirectoryEntryParams = $ContextSplatting
            $DirectoryEntryParams.remove('ContextType')
            $DirectoryEntry = New-ADSIDirectoryEntry @DirectoryEntryParams

            # Principal Searcher
            $DirectorySearcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher
            $DirectorySearcher.SearchRoot = $DirectoryEntry

            $DirectorySearcher.Filter = "(&(objectCategory=user)$LDAPFilter)"
            #$DirectorySearcher.PropertiesToLoad.AddRange("'Enabled','SamAccountName','DistinguishedName','Sid','DistinguishedName'")

            if (-not$PSBoundParameters['NoResultLimit']) {
                Write-Warning -Message "Result is limited to 1000 entries, specify a specific number on the parameter SizeLimit or 0 to remove the limit"
            }
            else {
                # SizeLimit is useless, even if there is a$Searcher.GetUnderlyingSearcher().sizelimit=$SizeLimit
                # the server limit is kept
                $DirectorySearcher.PageSize = 10000
            }

            $DirectorySearcher.FindAll() | ForEach-Object -Process {
                [System.DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($Context, $_.Properties["distinguishedname"])
            }# Return UserPrincipale object
        }
        else {
            Write-Verbose -Message "Searcher"

            $UserPrincipal = New-Object -TypeName System.DirectoryServices.AccountManagement.UserPrincipal -ArgumentList $Context
            $Searcher = New-Object -TypeName System.DirectoryServices.AccountManagement.PrincipalSearcher
            $Searcher.QueryFilter = $UserPrincipal

            if (-not$PSBoundParameters['NoResultLimit']) {
                Write-Warning -Message "Result is limited to 1000 entries, specify a specific number on the parameter SizeLimit or 0 to remove the limit"
            }
            else {
                # SizeLimit is useless, even if there is a$Searcher.GetUnderlyingSearcher().sizelimit=$SizeLimit
                # the server limit is kept
                $Searcher.GetUnderlyingSearcher().pagesize = 10000

            }
            #$Searcher.GetUnderlyingSearcher().propertiestoload.AddRange("'Enabled','SamAccountName','DistinguishedName','Sid','DistinguishedName'")
            $Searcher.FindAll() # Return UserPrincipale
        }
    }
}


function Get-ADHash {
    $ADUsers = Get-ADSIUser -NoResultLimit
    $UserHash = @{ }
    $UserList = $ADUsers.where( { $_.UserPrincipalName })
    foreach ($User in $UserList) {
        $UserHash[$User.UserPrincipalName] = @{
            Department     = $User.GetUnderlyingObject() | Select-Object -ExpandProperty Department
            Enabled        = $User.Enabled
            SamAccountName = $User.SamAccountName
        }
    }
    $UserHash
}

$InstallSplat = @{
    Name        = 'ImportExcel'
    Scope       = 'CurrentUser'
    Force       = $true
    ErrorAction = 'stop'
    Confirm     = $false
}

try {
    Install-Module @InstallSplat
}
catch {
    $_.Exception.Message
}



function Get-Answer {
    $Answer = Read-Host "Connect to Exchange Server? (Y/N)"
    if ($Answer -ne "Y" -and $Answer -ne "N") {
        Get-Answer
    }
    if ($Answer -eq "Y") {
        $ServerName = Read-Host "Type the name of the Exchange Server and hit enter"
        Connect-Exchange -Server $ServerName
    }

}


Get-Answer
Get-MailboxMoveOnPremisesMailboxReport -ReportPath ([Environment]::GetFolderPath("Desktop")) -Verbose





#######


