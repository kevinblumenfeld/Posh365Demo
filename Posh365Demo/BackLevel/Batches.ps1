
if (-not (Get-Module ActiveDirectory -listavailable)) {
    Write-Host "Please run from a computer with AD module" -ForegroundColor Red
    break
}
$Exec = Get-ExecutionPolicy
if ($Exec -eq 'Restricted') {
    Set-ExecutionPolicy -ExecutionPolicy Bypass -Force -Confirm:$false -ErrorAction SilentlyContinue
}

Function Get-MailboxMoveOnPremisesMailboxReport {
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
            'BatchName', 'DisplayName', 'OrganizationalUnit', 'CompleteBatchDate'
            'CompleteBatchTimePT', 'MailboxGB', 'ArchiveGB', 'DeletedGB', 'TotalGB'
            'LastLogonTime', 'ItemCount', 'UserPrincipalName', 'PrimarySmtpAddress'
            'AddressBookPolicy', 'RetentionPolicy', 'AccountDisabled', 'Alias'
            'Database', 'OU', 'Office', 'RecipientTypeDetails', 'UMEnabled'
            'ForwardingAddress', 'ForwardingRecipientType', 'DeliverToMailboxAndForward'
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
            ErrorAction             = 'SilentlyContinue'
        }
        $BatchesFile | Where-Object { $_ } | ForEach-Object { Import-Csv $_ | Export-Excel @ExcelSplat }
    }
}

Function Get-MailboxMoveOnPremisesReportHelper {
    [CmdletBinding()]
    param (

    )
    end {
        $RecHash = Get-MailboxMoveRecipientHash
        $MailboxList = Get-Mailbox -ResultSize Unlimited
        foreach ($Mailbox in $MailboxList) {
            Write-Verbose "Mailbox`t$($Mailbox.DisplayName)"
            $Statistic = $Mailbox | Get-ExchangeMailboxStatistics
            $PSHash = @{
                BatchName            = ''
                DisplayName          = $Mailbox.DisplayName
                OrganizationalUnit   = $Mailbox.OrganizationalUnit
                CompleteBatchDate    = ''
                CompleteBatchTimePT  = ''
                MailboxGB            = $Statistic.MailboxGB
                ArchiveGB            = $Statistic.ArchiveGB
                DeletedGB            = $Statistic.DeletedGB
                TotalGB              = $Statistic.TotalGB
                LastLogonTime        = $Statistic.LastLogonTime
                ItemCount            = $Statistic.ItemCount
                UserPrincipalName    = $Mailbox.UserPrincipalName
                PrimarySmtpAddress   = $Mailbox.PrimarySmtpAddress
                AddressBookPolicy    = $Mailbox.AddressBookPolicy
                RetentionPolicy      = $Mailbox.RetentionPolicy
                AccountDisabled      = $Mailbox.AccountDisabled
                Alias                = $Mailbox.Alias
                Database             = $Mailbox.Database
                OU                   = ($Mailbox.DistinguishedName -replace '^.+?,(?=(OU|CN)=)')
                Office               = $Mailbox.Office
                RecipientTypeDetails = $Mailbox.RecipientTypeDetails
                UMEnabled            = $Mailbox.UMEnabled
            }
            if ($Mailbox.ForwardingAddress) {
                $Distinguished = Convert-CanonicalToDistinguished -CanonicalName $Mailbox.ForwardingAddress
                $PSHash.Add('ForwardingAddress', $RecHash.$Distinguished.PrimarySmtpAddress)
                $PSHash.Add('ForwardingRecipientType', $RecHash.$Distinguished.RecipientTypeDetails)
                $PSHash.Add('DeliverToMailboxAndForward', $Mailbox.DeliverToMailboxAndForward)
            }
            else {
                $PSHash.Add('ForwardingAddress', '')
                $PSHash.Add('ForwardingRecipientType', '')
                $PSHash.Add('DeliverToMailboxAndForward', '')
            }
            New-Object -TypeName PSObject -Property $PSHash
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

Function Get-MailboxMoveRecipientHash {
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
        [System.Management.Automation.PSCredential]$Credential | Export-CliXml -Path $CredFile
        [System.Management.Automation.PSCredential]$Credential = Import-CliXml -Path $CredFile
    }
    else {
        [System.Management.Automation.PSCredential]$Credential = Import-CliXml -Path $CredFile
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


$InstallSplat = @{
    Name        = 'ImportExcel'
    Scope       = 'CurrentUser'
    Force       = $true
    ErrorAction = 'SilentlyContinue'
    Confirm     = $false
}

Install-Module @InstallSplat

function Get-Answer {
    $Answer = Read-Host "Connect to Exchange Server? (Y/N)"
    if ($Answer -ne "Y" -and $Answer -ne "N") {
        Get-Answer
    }
    if ($Answer -eq "Y") {
        $ServerName = Read-Host "Type then name of the Exchange Server and hit enter"
        Connect-Exchange -Server $ServerName
    }

}
Get-Answer
Get-MailboxMoveOnPremisesMailboxReport -ReportPath ([Environment]::GetFolderPath("Desktop")) -Verbose





#### END  ###
