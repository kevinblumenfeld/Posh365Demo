
if (-not (Get-Module ActiveDirectory -listavailable)) {
    Write-Host "Please run from a computer with AD module" -ForegroundColor Red
    break
}

Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force -ErrorAction SilentlyContinue
Import-Module ActiveDirectory -force

function Get-MailboxMoveOnPremisesPermissionReport {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $ReportPath,

        [Parameter()]
        [switch]
        $SkipSendAs,

        [Parameter()]
        [switch]
        $SkipSendOnBehalf,

        [Parameter()]
        [switch]
        $SkipFullAccess,

        [Parameter()]
        [switch]
        $SkipFolderPerms
    )
    end {
        New-Item -ItemType Directory -Path $ReportPath -ErrorAction SilentlyContinue

        Write-Verbose "Caching hashtable. msExchRecipientTypeDetails numerical value as Key and Value of human readable"
        $ADHashType = Get-ADHashType

        Write-Verbose "Caching hashtable. msExchRecipientDisplayType numerical value as Key and Value of human readable"
        $ADHashDisplay = Get-ADHashDisplay

        $DelegateSplat = @{
            SkipFullAccess   = $SkipFullAccess
            SkipSendOnBehalf = $SkipSendOnBehalf
            SkipSendAs       = $SkipSendAs
            ADHashType       = $ADHashType
            ADHashDisplay    = $ADHashDisplay
            ErrorAction      = 'SilentlyContinue'
        }
        if ($DelegateSplat.Values -contains $false) {
            try {
                Import-Module ActiveDirectory -ErrorAction Stop -Verbose:$false
            }
            catch {
                Write-Host "This module depends on the ActiveDirectory module."
                Write-Host "Please download and install from https://www.microsoft.com/en-us/download/details.aspx?id=45520"
                Write-Host "or run Connect-Exchange from a server with the Active Directory Module installed"
                throw
            }
        }
        $DomainNameHash = Get-DomainNameHash
        Write-Verbose "Importing Active Directory Users and Groups that have at least one proxy address"

        $ADUserList = Get-ADUsersAndGroups -DomainNameHash $DomainNameHash
        $UserGroupHash = @{ }
        $ADUserList | ForEach-Object { $usergrouphash.Add($_.ObjectGuid, @{
                    'PrimarySmtpAddress' = $_.PrimarySmtpAddress
                    'DisplayName'        = $_.DisplayName
                    'UserPrincipalName'  = $_.UserPrincipalName
                }) }

        $GroupMemberHash = Get-ADGroupMemberHash -DomainNameHash $DomainNameHash -UserGroupHash $UserGroupHash

        Write-Verbose "Retrieving all Exchange Mailboxes"
        $MailboxList = Get-Mailbox -ResultSize unlimited
        if ($DelegateSplat.Values -contains $false) {
            $DelegateSplat.Add('MailboxList', $MailboxList)
            $DelegateSplat.Add('ADUserList', $ADUserList)
            $DelegateSplat.Add('UserGroupHash', $UserGroupHash)
            $DelegateSplat.Add('GroupMemberHash', $GroupMemberHash)
            Get-MailboxMoveMailboxPermission @DelegateSplat | Export-Csv (Join-Path $ReportPath 'MailboxPermissions.csv') -NoTypeInformation -Encoding UTF8
            $MailboxFile = Join-Path $ReportPath 'MailboxPermissions.csv'
        }
        if (-not $SkipFolderPerms) {
            $FolderPermSplat = @{
                MailboxList   = $MailboxList
                ADUserList    = $ADUserList
                ADHashType    = $ADHashType
                ADHashDisplay = $ADHashDisplay
                ErrorAction   = 'SilentlyContinue'
            }
            Get-MailboxMoveFolderPermission @FolderPermSplat | Export-Csv (Join-Path $ReportPath 'FolderPermissions.csv') -NoTypeInformation -Encoding UTF8
            $FolderFile = Join-Path $ReportPath 'FolderPermissions.csv'
        }
        $ExcelSplat = @{
            Path                    = (Join-Path $ReportPath 'Permissions.xlsx')
            TableStyle              = 'Medium2'
            FreezeTopRowFirstColumn = $true
            AutoSize                = $true
            BoldTopRow              = $true
            ClearSheet              = $true
            ErrorAction             = 'SilentlyContinue'
        }
        $MailboxFile, $FolderFile | Where-Object { $_ } | ForEach-Object { Import-Csv $_ | Export-Excel @ExcelSplat -WorksheetName ($_ -replace '.+\\|permissions\.csv') }
    }
}

function Get-ADHashType {
    param (

    )
    end {
        $TypeDetails = @{
            '1'            = 'UserMailbox'
            '2'            = 'LinkedMailbox'
            '4'            = 'SharedMailbox'
            '8'            = 'LegacyMailbox'
            '16'           = 'RoomMailbox'
            '32'           = 'EquipmentMailbox'
            '64'           = 'MailContact'
            '128'          = 'MailEnabledUser'
            '256'          = 'MailEnabledUniversalDistributionGroup'
            '512'          = 'MailEnabledNonUniversalDistributionGroup'
            '1024'         = 'MailEnabledUniversalSecurityGroup'
            '2048'         = 'DynamicDistributionGroup'
            '4096'         = 'MailEnabledPublicFolder'
            '8192'         = 'SystemAttendantMailbox'
            '16384'        = 'MailboxDatabaseMailbox'
            '32768'        = 'AcrossForestMailContact'
            '65536'        = 'User'
            '131072'       = 'Contact'
            '262144'       = 'UniversalDistributionGroup'
            '524288'       = 'UniversalSecurityGroup'
            '1048576'      = 'Non-UniversalGroup'
            '2097152'      = 'DisabledUser'
            '4194304'      = 'MicrosoftExchange'
            '8388608'      = 'ArbitrationMailbox'
            '16777216'     = 'MailboxPlan'
            '33554432'     = 'LinkedUser'
            '268435456'    = 'RoomList'
            '536870912'    = 'DiscoverMailbox'
            '1073741824'   = 'RoleGroup'
            '2147483648'   = 'RemoteUserMailbox'
            '8589934592'   = 'RemoteRoomMailbox'
            '17173869184'  = 'RemoteEquipmentMailbox'
            '34359738368'  = 'RemoteSharedMailbox'
            '137438953472' = 'TeamMailbox'
        }
        $TypeDetails
    }
}

function Get-ADHashDisplay {
    param (

    )
    end {
        $Display = @{
            '0'           = 'MailboxUser'
            '1'           = 'DistributionGroup'
            '2'           = 'PublicFolder'
            '3'           = 'DynamicDistributionGroup'
            '4'           = 'Organization'
            '5'           = 'PrivateDistributionList'
            '6'           = 'RemoteMailUser'
            '7'           = 'ConferenceRoomMailbox'
            '8'           = 'EquipmentMailbox'
            '10'          = 'ArbitrationMailbox'
            '11'          = 'MailboxPlan'
            '12'          = 'LinkedUser'
            '15'          = 'RoomList'
            '1073741833'  = 'SecurityDistributionGroup'
            '1073741824'  = 'ACLableMailboxUser'
            '1073741830'  = 'ACLableRemoteMailUser'
            '-2147481343' = 'SyncedUSGasUDG'
            '-1073739511' = 'SyncedUSGasUSG'
            '-2147481338' = 'SyncedUSGasContact'
            '-1073739514' = 'ACLableSyncedUSGasContact'
            '-2147482874' = 'SyncedDynamicDistributionGroup'
            '-1073741818' = 'ACLableSyncedMailboxUser'
            '-2147483642' = 'SyncedMailboxUser'
            '-2147481850' = 'SyncedConferenceRoomMailbox'
            '-2147481594' = 'SyncedEquipmentMailbox'
            '-2147482106' = 'SyncedRemoteMailUser'
            '-1073740282' = 'ACLableSyncedRemoteMailUser'
            '-2147483130' = 'SyncedPublicFolder'

        }
        $Display
    }
}

function Get-DomainNameHash {

    param (

    )
    end {
        $DomainNameHash = @{ }

        $DomainList = ([System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().domains) | Select -ExpandProperty Name
        foreach ($Domain in $DomainList) {
            $DomainNameHash[$Domain] = (ConvertTo-NetBios -domain $Domain)
        }
        $DomainNameHash
    }
}

function Get-ADUsersAndGroups {
    param (
        [Parameter()]
        [hashtable] $DomainNameHash
    )
    try {
        import-module activedirectory -ErrorAction Stop -Verbose:$false
    }
    catch {
        Write-Host "This module depends on the ActiveDirectory module."
        Write-Host "Please download and install from https://www.microsoft.com/en-us/download/details.aspx?id=45520"
        throw
    }

    # Find writable Global Catalog
    $context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Forest')
    $dc = ([System.DirectoryServices.ActiveDirectory.GlobalCatalog]::FindOne($context, [System.DirectoryServices.ActiveDirectory.LocatorOptions]'ForceRediscovery, WriteableRequired')).name

    $Selectproperties = @(
        'DisplayName', 'UserPrincipalName', 'distinguishedname', 'SamAccountName', 'ProxyAddresses'
        'canonicalname', 'mail', 'Objectguid', 'msExchRecipientTypeDetails'
        'msExchRecipientDisplayType', 'objectClass'
    )
    $CalculatedProps = @(
        @{n = "logon"; e = { ($DomainNameHash.($_.distinguishedname -replace '^.+?DC=' -replace ',DC=', '.')) + "\" + $_.samaccountname } },
        @{n = "PrimarySMTPAddress" ; e = { ( $_.proxyAddresses | Where-Object { $_ -cmatch "SMTP:*" }).Substring(5) } }
    )
    Get-ADUser -filter * -server ($dc + ":3268") -SearchBase (Get-ADRootDSE).rootdomainnamingcontext -SearchScope Subtree -Properties $SelectProperties |
    Select-Object ($Selectproperties + $CalculatedProps)
    Get-ADGroup -filter * -server ($dc + ":3268") -SearchBase (Get-ADRootDSE).rootdomainnamingcontext -SearchScope Subtree -Properties $SelectProperties |
    Select-Object ($Selectproperties + $CalculatedProps)
}

function Get-MailboxMoveMailboxPermission {
    [CmdletBinding()]
    param (
        [Parameter()]
        [switch]
        $SkipSendAs,

        [Parameter()]
        [switch]
        $SkipSendOnBehalf,

        [Parameter()]
        [switch]
        $SkipFullAccess,

        [Parameter(Mandatory = $true)]
        $MailboxList,

        [Parameter(Mandatory = $true)]
        $ADUserList,

        [parameter()]
        [hashtable]
        $ADHashType,

        [parameter()]
        [hashtable]
        $ADHashDisplay,

        [parameter()]
        [hashtable]
        $UserGroupHash,

        [parameter()]
        [hashtable]
        $GroupMemberHash
    )
    end {
        Write-Verbose "Caching hashtable. LogonName as Key and Values of DisplayName & UPN"
        $ADHash = $ADUserList | Get-ADHash

        Write-Verbose "Caching hashtable. DN as Key and Values of DisplayName, UPN & LogonName"
        $ADHashDN = $ADUserList | Get-ADHashDN

        Write-Verbose "Caching hashtable. CN as Key and Values of DisplayName, UPN & LogonName"
        $ADHashCN = $ADUserList | Get-ADHashCN

        $MailboxDN = $MailboxList | Select-Object -expandproperty distinguishedname

        $PermSelect = @(
            'Object', 'UserPrincipalName', 'PrimarySMTPAddress', 'Granted', 'GrantedUPN'
            'GrantedSMTP', 'Checking', 'TypeDetails', 'DisplayType', 'Permission'
        )
        $ParamSplat = @{
            ADHashDN        = $ADHashDN
            ADHash          = $ADHash
            ADHashType      = $ADHashType
            ADHashDisplay   = $ADHashDisplay
            UserGroupHash   = $UserGroupHash
            GroupMemberHash = $GroupMemberHash
        }
        $ParamSOBSplat = @{
            ADHashCN        = $ADHashCN
            ADHashDN        = $ADHashDN
            ADHashType      = $ADHashType
            ADHashDisplay   = $ADHashDisplay
            UserGroupHash   = $UserGroupHash
            GroupMemberHash = $GroupMemberHash
        }
        if (-not $SkipSendAs) {
            Write-Verbose "Getting SendAs permissions for each mailbox and writing to file"
            $MailboxDN | Get-SendAsPerms @ParamSplat |
            Select-Object $PermSelect
        }
        if (-not $SkipSendOnBehalf) {
            Write-Verbose "Getting SendOnBehalf permissions for each mailbox and writing to file"
            ($MailboxList | Where-Object { $_.GrantSendOnBehalfTo }) | Get-SendOnBehalfPerms @ParamSOBSplat |
            Select-Object $PermSelect
        }
        if (-not $SkipFullAccess) {
            Write-Verbose "Getting FullAccess permissions for each mailbox and writing to file"
            $MailboxDN | Get-FullAccessPerms @ParamSplat |
            Select-Object $PermSelect
        }
    }
}

Function Get-MailboxMoveFolderPermission {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        $MailboxList,

        [Parameter(Mandatory = $true)]
        $ADUserList,

        [parameter()]
        [hashtable]
        $ADHashType,

        [parameter()]
        [hashtable]
        $ADHashDisplay
    )
    end {
        $FolderSelect = @(
            'Object', 'UserPrincipalName', 'PrimarySMTPAddress', 'Folder', 'AccessRights'
            'Granted', 'GrantedUPN', 'GrantedSMTP', 'TypeDetails', 'DisplayType'
        )
        Write-Verbose "Caching hashtable. DisplayName as Key and Values of UPN, PrimarySMTP, msExchRecipientTypeDetails & msExchRecipientDisplayType"
        $ADHashDisplayName = $ADUserList | Get-ADHashDisplayName -erroraction silentlycontinue

        $FolderPermSplat = @{
            ADHashDisplayName = $ADHashDisplayName
            ADHashType        = $ADHashType
            ADHashDisplay     = $ADHashDisplay
            UserGroupHash     = $UserGroupHash
            GroupMemberHash   = $GroupMemberHash

        }
        Write-Verbose "Getting Folder Permissions for each mailbox and writing to file"
        $MailboxList | Get-MailboxFolderPerms @FolderPermSplat | Select-Object $FolderSelect
    }
}


function Get-MailboxFolderPerms {
    [CmdletBinding()]
    Param (
        [parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        $MailboxList,

        [parameter()]
        [hashtable]
        $ADHashDisplayName,

        [parameter()]
        [hashtable]
        $ADHashType,

        [parameter()]
        [hashtable]
        $ADHashDisplay,

        [parameter()]
        [hashtable]
        $UserGroupHash,

        [parameter()]
        [hashtable]
        $GroupMemberHash
    )
    begin {

    }
    process {
        foreach ($Mailbox in $MailboxList) {
            Write-Verbose "Inspecting: `t $($Mailbox.UserPrincipalName)"
            $StatSplat = @{
                Identity    = $Mailbox.UserPrincipalName
                ErrorAction = 'SilentlyContinue'
            }
            $Calendar = (($Mailbox.SamAccountName) + ":\" + (Get-MailboxFolderStatistics @StatSplat -FolderScope Calendar | Select-Object -First 1).Name)
            $Inbox = (($Mailbox.SamAccountName) + ":\" + (Get-MailboxFolderStatistics @StatSplat -FolderScope Inbox | Select-Object -First 1).Name)
            $SentItems = (($Mailbox.SamAccountName) + ":\" + (Get-MailboxFolderStatistics @StatSplat -FolderScope SentItems | Select-Object -First 1).Name)
            $Contacts = (($Mailbox.SamAccountName) + ":\" + (Get-MailboxFolderStatistics @StatSplat -FolderScope Contacts | Select-Object -First 1).Name)
            $CalAccessList = Get-MailboxFolderPermission $Calendar | Where-Object {
                $_.User -notmatch 'Default' -and
                $_.User -notmatch 'Anonymous' -and
                $_.User -notlike 'NT User:*' -and
                $_.AccessRights -notmatch 'None'
            }
            If ($CalAccessList) {
                Foreach ($CalAccess in $CalAccessList) {
                    $Logon = $ADHashDisplayName[$CalAccess.User].logon
                    $DisplayType = $ADHashDisplayName[$CalAccess.User].msExchRecipientDisplayType
                    if ($GroupMemberHash[$Logon].Members -and $ADHashDisplay["$DisplayType"] -match 'group') {
                        foreach ($Member in @($GroupMemberHash.$Logon.Members)) {
                            Write-Verbose "  Member:`t$Member"
                            New-Object -TypeName psobject -property @{
                                Object             = $Mailbox.DisplayName
                                UserPrincipalName  = $Mailbox.UserPrincipalName
                                PrimarySMTPAddress = $Mailbox.PrimarySMTPAddress
                                Folder             = 'CALENDAR'
                                AccessRights       = ($CalAccess.AccessRights) -join ','
                                Granted            = $UserGroupHash[$Member].DisplayName
                                GrantedUPN         = $UserGroupHash[$Member].UserPrincipalName
                                GrantedSMTP        = $UserGroupHash[$Member].PrimarySMTPAddress
                                TypeDetails        = "GroupMember"
                                DisplayType        = $ADHashDisplay."$($ADHashDisplayName."$($CalAccess.User)".msExchRecipientDisplayType)"
                            }
                        }
                    }
                    elseif ( $ADHashDisplayName[$CalAccess.User].objectClass -notmatch 'group') {
                        New-Object -TypeName psobject -property @{
                            Object             = $Mailbox.DisplayName
                            UserPrincipalName  = $Mailbox.UserPrincipalName
                            PrimarySMTPAddress = $Mailbox.PrimarySMTPAddress
                            Folder             = 'CALENDAR'
                            AccessRights       = ($CalAccess.AccessRights) -join ','
                            Granted            = $CalAccess.User
                            GrantedUPN         = $ADHashDisplayName."$($CalAccess.User)".UserPrincipalName
                            GrantedSMTP        = $ADHashDisplayName."$($CalAccess.User)".PrimarySMTPAddress
                            TypeDetails        = $ADHashType."$($ADHashDisplayName."$($CalAccess.User)".msExchRecipientTypeDetails)"
                            DisplayType        = $ADHashDisplay."$($ADHashDisplayName."$($CalAccess.User)".msExchRecipientDisplayType)"
                        }
                    }
                }
            }
            $InboxAccessList = Get-MailboxFolderPermission $Inbox | Where-Object {
                $_.User -notmatch 'Default' -and
                $_.User -notmatch 'Anonymous' -and
                $_.User -notlike 'NT User:*' -and
                $_.AccessRights -notmatch 'None'
            }
            If ($InboxAccessList) {
                Foreach ($InboxAccess in $InboxAccessList) {
                    $Logon = $ADHashDisplayName[$InboxAccess.User].logon
                    $DisplayType = $ADHashDisplayName[$InboxAccess.User].msExchRecipientDisplayType
                    if ($GroupMemberHash[$Logon].Members -and $ADHashDisplay["$DisplayType"] -match 'group') {
                        foreach ($Member in @($GroupMemberHash.$Logon.Members)) {
                            Write-Verbose "  Member:`t$Member"
                            New-Object -TypeName psobject -property @{
                                Object             = $Mailbox.DisplayName
                                UserPrincipalName  = $Mailbox.UserPrincipalName
                                PrimarySMTPAddress = $Mailbox.PrimarySMTPAddress
                                Folder             = 'INBOX'
                                AccessRights       = ($InboxAccess.AccessRights) -join ','
                                Granted            = $UserGroupHash[$Member].DisplayName
                                GrantedUPN         = $UserGroupHash[$Member].UserPrincipalName
                                GrantedSMTP        = $UserGroupHash[$Member].PrimarySMTPAddress
                                TypeDetails        = "GroupMember"
                                DisplayType        = $ADHashDisplay."$($ADHashDisplayName."$($InboxAccess.User)".msExchRecipientDisplayType)"
                            }
                        }
                    }
                    elseif ( $ADHashDisplayName[$InboxAccess.User].objectClass -notmatch 'group') {
                        New-Object -TypeName psobject -property @{
                            Object             = $Mailbox.DisplayName
                            UserPrincipalName  = $Mailbox.UserPrincipalName
                            PrimarySMTPAddress = $Mailbox.PrimarySMTPAddress
                            Folder             = 'INBOX'
                            AccessRights       = ($InboxAccess.AccessRights) -join ','
                            Granted            = $InboxAccess.User
                            GrantedUPN         = $ADHashDisplayName."$($InboxAccess.User)".UserPrincipalName
                            GrantedSMTP        = $ADHashDisplayName."$($InboxAccess.User)".PrimarySMTPAddress
                            TypeDetails        = $ADHashType."$($ADHashDisplayName."$($InboxAccess.User)".msExchRecipientTypeDetails)"
                            DisplayType        = $ADHashDisplay."$($ADHashDisplayName."$($InboxAccess.User)".msExchRecipientDisplayType)"
                        }
                    }
                }
            }
            $SentAccessList = Get-MailboxFolderPermission $SentItems | Where-Object {
                $_.User -notmatch 'Default' -and
                $_.User -notmatch 'Anonymous' -and
                $_.User -notlike 'NT User:*' -and
                $_.AccessRights -notmatch 'None'
            }
            If ($SentAccessList) {
                Foreach ($SentAccess in $SentAccessList) {
                    $Logon = $ADHashDisplayName[$SentAccess.User].logon
                    $DisplayType = $ADHashDisplayName[$SentAccess.User].msExchRecipientDisplayType
                    if ($GroupMemberHash[$Logon].Members -and $ADHashDisplay["$DisplayType"] -match 'group') {
                        foreach ($Member in @($GroupMemberHash.$Logon.Members)) {
                            Write-Verbose "  Member:`t$Member"
                            New-Object -TypeName psobject -property @{
                                Object             = $Mailbox.DisplayName
                                UserPrincipalName  = $Mailbox.UserPrincipalName
                                PrimarySMTPAddress = $Mailbox.PrimarySMTPAddress
                                Folder             = 'SENTITEMS'
                                AccessRights       = ($SentAccess.AccessRights) -join ','
                                Granted            = $UserGroupHash[$Member].DisplayName
                                GrantedUPN         = $UserGroupHash[$Member].UserPrincipalName
                                GrantedSMTP        = $UserGroupHash[$Member].PrimarySMTPAddress
                                TypeDetails        = "GroupMember"
                                DisplayType        = $ADHashDisplay."$($ADHashDisplayName."$($SentAccess.User)".msExchRecipientDisplayType)"
                            }
                        }
                    }
                    elseif ( $ADHashDisplayName[$SentAccess.User].objectClass -notmatch 'group') {
                        New-Object -TypeName psobject -property @{
                            Object             = $Mailbox.DisplayName
                            UserPrincipalName  = $Mailbox.UserPrincipalName
                            PrimarySMTPAddress = $Mailbox.PrimarySMTPAddress
                            Folder             = 'SENTITEMS'
                            AccessRights       = ($SentAccess.AccessRights) -join ','
                            Granted            = $SentAccess.User
                            GrantedUPN         = $ADHashDisplayName."$($SentAccess.User)".UserPrincipalName
                            GrantedSMTP        = $ADHashDisplayName."$($SentAccess.User)".PrimarySMTPAddress
                            TypeDetails        = $ADHashType."$($ADHashDisplayName."$($SentAccess.User)".msExchRecipientTypeDetails)"
                            DisplayType        = $ADHashDisplay."$($ADHashDisplayName."$($SentAccess.User)".msExchRecipientDisplayType)"
                        }
                    }
                }
            }
            $ContactsAccessList = Get-MailboxFolderPermission $Contacts | Where-Object {
                $_.User -notmatch 'Default' -and
                $_.User -notmatch 'Anonymous' -and
                $_.User -notlike 'NT User:*' -and
                $_.AccessRights -notmatch 'None'
            }
            If ($ContactsAccessList) {
                Foreach ($ContactsAccess in $ContactsAccessList) {
                    $Logon = $ADHashDisplayName[$ContactsAccess.User].logon
                    $DisplayType = $ADHashDisplayName[$ContactsAccess.User].msExchRecipientDisplayType
                    if ($GroupMemberHash[$Logon].Members -and $ADHashDisplay["$DisplayType"] -match 'group') {
                        foreach ($Member in @($GroupMemberHash.$Logon.Members)) {
                            Write-Verbose "  Member:`t$Member"
                            New-Object -TypeName psobject -property @{
                                Object             = $Mailbox.DisplayName
                                UserPrincipalName  = $Mailbox.UserPrincipalName
                                PrimarySMTPAddress = $Mailbox.PrimarySMTPAddress
                                Folder             = 'CONTACTS'
                                AccessRights       = ($ContactsAccess.AccessRights) -join ','
                                Granted            = $UserGroupHash[$Member].DisplayName
                                GrantedUPN         = $UserGroupHash[$Member].UserPrincipalName
                                GrantedSMTP        = $UserGroupHash[$Member].PrimarySMTPAddress
                                TypeDetails        = "GroupMember"
                                DisplayType        = $ADHashDisplay."$($ADHashDisplayName."$($ContactsAccess.User)".msExchRecipientDisplayType)"
                            }
                        }
                    }
                    elseif ( $ADHashDisplayName[$ContactsAccess.User].objectClass -notmatch 'group') {
                        New-Object -TypeName psobject -property @{
                            Object             = $Mailbox.DisplayName
                            UserPrincipalName  = $Mailbox.UserPrincipalName
                            PrimarySMTPAddress = $Mailbox.PrimarySMTPAddress
                            Folder             = 'CONTACTS'
                            AccessRights       = ($ContactsAccess.AccessRights) -join ','
                            Granted            = $ContactsAccess.User
                            GrantedUPN         = $ADHashDisplayName."$($ContactsAccess.User)".UserPrincipalName
                            GrantedSMTP        = $ADHashDisplayName."$($ContactsAccess.User)".PrimarySMTPAddress
                            TypeDetails        = $ADHashType."$($ADHashDisplayName."$($ContactsAccess.User)".msExchRecipientTypeDetails)"
                            DisplayType        = $ADHashDisplay."$($ADHashDisplayName."$($ContactsAccess.User)".msExchRecipientDisplayType)"
                        }
                    }
                }
            }
        }
    }
    end {

    }
}

function Get-ADHashDisplayName {

    param (
        [parameter(ValueFromPipeline = $true)]
        $ADUserList
    )
    begin {
        $ADHashDisplayName = @{ }
    }
    process {
        foreach ($ADUser in $ADUserList) {
            $ADHashDisplayName[$ADUser.DisplayName] = @{
                UserPrincipalName          = $ADUser.UserPrincipalName
                PrimarySMTPAddress         = $ADUser.PrimarySMTPAddress
                msExchRecipientTypeDetails = $ADUser.msExchRecipientTypeDetails
                msExchRecipientDisplayType = $ADUser.msExchRecipientDisplayType
                Logon                      = $ADUser.Logon
                objectClass                = $ADUser.objectClass
            }
        }
    }
    end {
        $ADHashDisplayName
    }
}

function Get-ADHashCN {
    param (
        [parameter(ValueFromPipeline = $true)]
        $ADUserList
    )
    begin {
        $ADHashCN = @{ }
    }
    process {
        foreach ($ADUser in $ADUserList) {
            $ADHashCN[$ADUser.CanonicalName] = @{
                DisplayName                = $ADUser.DisplayName
                UserPrincipalName          = $ADUser.UserPrincipalName
                Logon                      = $ADUser.logon
                PrimarySMTPAddress         = $ADUser.PrimarySMTPAddress
                msExchRecipientTypeDetails = $ADUser.msExchRecipientTypeDetails
                msExchRecipientDisplayType = $ADUser.msExchRecipientDisplayType
            }
        }
    }
    end {
        $ADHashCN
    }
}

function Get-ADHashDN {
    param (
        [parameter(ValueFromPipeline = $true)]
        $MailboxList
    )
    begin {
        $ADHashDN = @{ }
    }
    process {
        foreach ($Mailbox in $MailboxList) {
            $ADHashDN[$Mailbox.DistinguishedName] = @{
                DisplayName        = $Mailbox.DisplayName
                UserPrincipalName  = $Mailbox.UserPrincipalName
                Logon              = $Mailbox.logon
                PrimarySMTPAddress = $Mailbox.PrimarySMTPAddress
            }
        }
    }
    end {
        $ADHashDN
    }
}
#
function Get-ADHash {
    param (
        [parameter(ValueFromPipeline = $true)]
        $ADUserList
    )
    begin {
        $ADHash = @{ }
    }
    process {
        foreach ($ADUser in $ADUserList) {
            $ADHash[$ADUser.logon] = @{
                DisplayName                = $ADUser.DisplayName
                UserPrincipalName          = $ADUser.UserPrincipalName
                PrimarySMTPAddress         = $ADUser.PrimarySMTPAddress
                msExchRecipientTypeDetails = $ADUser.msExchRecipientTypeDetails
                msExchRecipientDisplayType = $ADUser.msExchRecipientDisplayType
                Objectguid                 = $ADUser.Objectguid
                objectClass                = $ADUser.objectClass
            }
        }
    }
    end {
        $ADHash
    }
}
function Get-ADGroupMemberHash {
    param (
        [Parameter()]
        [hashtable] $DomainNameHash,

        [Parameter()]
        [hashtable] $UserGroupHash
    )
    $context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Forest')
    $dc = ([System.DirectoryServices.ActiveDirectory.GlobalCatalog]::FindOne($context, [System.DirectoryServices.ActiveDirectory.LocatorOptions]'ForceRediscovery, WriteableRequired')).name
    $GroupMemberHash = @{ }
    $GroupParams = @{
        LDAPFilter  = "(!(SamAccountName=Domain Computers))"
        Server      = ($dc + ':3268')
        SearchBase  = (Get-ADRootDSE).rootdomainnamingcontext
        SearchScope = 'Subtree'
    }
    Get-ADGroup @GroupParams | ForEach-Object {
        write-host "Caching Members:`t$($_.Name)" -ForegroundColor Green
        $GroupMemberHash.Add( ($DomainNameHash.($_.distinguishedname -replace '^.+?DC=' -replace ',DC=', '.')) + "\" + $_.samaccountname, @{
                SID     = $_.SID
                MEMBERS = @(Get-ADGroupMember -Identity $_.SID -Recursive) -ne '' | foreach-object { $_.ObjectGuid }
            } )
    }
    $GroupMemberHash
}

function ConvertTo-NetBios {

    Param(
        $domainName
    )

    $root = [adsi] "LDAP://$domainname/RootDSE"
    $configContext = $root.Properties["configurationNamingContext"][0]
    $searchr = [adsi] "LDAP://cn=Partitions,$configContext"

    $search = new-object System.DirectoryServices.DirectorySearcher
    $search.SearchRoot = $searchr
    $search.SearchScope = [System.DirectoryServices.SearchScope] "OneLevel"
    $search.filter = "(&(objectcategory=Crossref)(dnsRoot=$domainName)(netBIOSName=*))"

    $result = $search.Findone()

    if ($result) {

        $result.Properties["netbiosname"]
    }

}
function Get-SendAsPerms {
    [CmdletBinding()]
    Param (
        [parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        $DistinguishedName,

        [parameter()]
        [hashtable]
        $ADHashDN,

        [parameter()]
        [hashtable]
        $ADHash,

        [parameter()]
        [hashtable]
        $ADHashType,

        [parameter()]
        [hashtable]
        $ADHashDisplay,

        [parameter()]
        [hashtable]
        $UserGroupHash,

        [parameter()]
        [hashtable]
        $GroupMemberHash
    )
    process {
        foreach ($ADUser in $DistinguishedName) {
            Write-Verbose "Inspecting:`t $ADUser"
            Get-ADPermission $ADUser | Where-Object {
                $_.ExtendedRights -like "*Send-As*" -and
                ($_.IsInherited -eq $false) -and
                !($_.User -like "NT AUTHORITY\SELF") -and
                !($_.User.tostring().startswith('S-1-5-21-')) -and
                !$_.Deny
            } | ForEach-Object {
                $HasPerm = $_.User
                if ($GroupMemberHash[$HasPerm].Members -and
                    $ADHashDisplay."$($ADHash["$HasPerm"].msExchRecipientDisplayType)" -match 'group') {
                    foreach ($Member in @($GroupMemberHash[$HasPerm].Members)) {
                        New-Object -TypeName psobject -property @{
                            Object             = $ADHashDN["$ADUser"].DisplayName
                            UserPrincipalName  = $ADHashDN["$ADUser"].UserPrincipalName
                            PrimarySMTPAddress = $ADHashDN["$ADUser"].PrimarySMTPAddress
                            Granted            = $UserGroupHash[$Member].DisplayName
                            GrantedUPN         = $UserGroupHash[$Member].UserPrincipalName
                            GrantedSMTP        = $UserGroupHash[$Member].PrimarySMTPAddress
                            Checking           = $HasPerm
                            TypeDetails        = "GroupMember"
                            DisplayType        = $ADHashDisplay."$($ADHash["$HasPerm"].msExchRecipientDisplayType)"
                            Permission         = "SendAs"
                        }
                    }
                }
                elseif ( $ADHash["$HasPerm"].objectClass -notmatch 'group') {
                    Write-Verbose "Has Send As:`t $($HasPerm)"
                    New-Object -TypeName psobject -property @{
                        Object             = $ADHashDN["$ADUser"].DisplayName
                        UserPrincipalName  = $ADHashDN["$ADUser"].UserPrincipalName
                        PrimarySMTPAddress = $ADHashDN["$ADUser"].PrimarySMTPAddress
                        Granted            = $ADHash["$($HasPerm)"].DisplayName
                        GrantedUPN         = $ADHash["$($HasPerm)"].UserPrincipalName
                        GrantedSMTP        = $ADHash["$($HasPerm)"].PrimarySMTPAddress
                        Checking           = $HasPerm
                        TypeDetails        = $ADHashType."$($ADHash["$($HasPerm)"].msExchRecipientTypeDetails)"
                        DisplayType        = $ADHashDisplay."$($ADHash["$($HasPerm)"].msExchRecipientDisplayType)"
                        Permission         = "SendAs"
                    }
                }
            }
        }
    }
    end {

    }
}

function Get-SendOnBehalfPerms {
    [CmdletBinding()]
    Param (
        [parameter(ValueFromPipeline = $true)]
        $MailboxList,

        [parameter()]
        [hashtable]
        $ADHashDN,

        [parameter()]
        [hashtable]
        $ADHashCN,

        [parameter()]
        [hashtable]
        $ADHashType,

        [parameter()]
        [hashtable]
        $ADHashDisplay,

        [parameter()]
        [hashtable]
        $UserGroupHash,

        [parameter()]
        [hashtable]
        $GroupMemberHash
    )
    begin {

    }
    process {
        foreach ($Mailbox in $MailboxList) {
            Write-Verbose "Inspecting: `t $Mailbox"

            foreach ($HasPerm in @($Mailbox.GrantSendOnBehalfTo)) {
                $Logon = $ADHashCN.$HasPerm.logon
                if ($GroupMemberHash.$Logon.Members -and
                    $ADHashDisplay."$($ADHashCN["$HasPerm"].msExchRecipientDisplayType)" -match 'group') {
                    foreach ($Member in @($GroupMemberHash.$Logon.Members)) {
                        Write-Verbose "  Member: `t $Member"
                        New-Object -TypeName psobject -property @{
                            Object             = $Mailbox.DisplayName
                            UserPrincipalName  = $Mailbox.UserPrincipalName
                            PrimarySMTPAddress = $Mailbox.PrimarySMTPAddress
                            Granted            = $UserGroupHash[$Member].DisplayName
                            GrantedUPN         = $UserGroupHash[$Member].UserPrincipalName
                            GrantedSMTP        = $UserGroupHash[$Member].PrimarySMTPAddress
                            Checking           = $ADHashCN.$HasPerm.DisplayName
                            TypeDetails        = "GroupMember"
                            DisplayType        = $ADHashDisplay."$($ADHashCN["$HasPerm"].msExchRecipientDisplayType)"
                            Permission         = "SendOnBehalf"
                        }
                    }
                }
                elseif ( $ADHash["$HasPerm"].objectClass -notmatch 'group') {
                    Write-Verbose "  CN: `t $HasPerm"
                    New-Object -TypeName psobject -property @{
                        Object             = $Mailbox.DisplayName
                        UserPrincipalName  = $Mailbox.UserPrincipalName
                        PrimarySMTPAddress = $Mailbox.PrimarySMTPAddress
                        Granted            = $ADHashCN["$HasPerm"].DisplayName
                        GrantedUPN         = $ADHashCN["$HasPerm"].UserPrincipalName
                        GrantedSMTP        = $ADHashCN["$HasPerm"].PrimarySMTPAddress
                        Checking           = $ADHashCN.$HasPerm.DisplayName
                        TypeDetails        = $ADHashType."$($ADHashCN["$HasPerm"].msExchRecipientTypeDetails)"
                        DisplayType        = $ADHashDisplay."$($ADHashCN["$HasPerm"].msExchRecipientDisplayType)"
                        Permission         = "SendOnBehalf"
                    }
                }
            }
        }
    }
    end {

    }
}

function Get-FullAccessPerms {

    [CmdletBinding()]
    Param (
        [parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        $ADUserList,

        [parameter()]
        [hashtable]
        $ADHashDN,

        [parameter()]
        [hashtable]
        $ADHash,

        [parameter()]
        [hashtable]
        $ADHashType,

        [parameter()]
        [hashtable]
        $ADHashDisplay,

        [parameter()]
        [hashtable]
        $UserGroupHash,

        [parameter()]
        [hashtable]
        $GroupMemberHash
    )
    begin {

    }
    process {
        foreach ($ADUser in $ADUserList) {
            Write-Verbose "Inspecting:`t $ADUser"
            Get-MailboxPermission $ADUser |
            Where-Object {
                $_.AccessRights -like "*FullAccess*" -and
                !$_.IsInherited -and !$_.user.tostring().startswith('S-1-5-21-') -and
                !$_.user.tostring().startswith('NT AUTHORITY\SELF') -and
                !$_.Deny
            } | ForEach-Object {
                $HasPerm = $_.User
                if ($GroupMemberHash[$HasPerm].Members -and
                    $ADHashDisplay."$($ADHash["$HasPerm"].msExchRecipientDisplayType)" -match 'group') {
                    foreach ($Member in @($GroupMemberHash[$HasPerm].Members)) {
                        New-Object -TypeName psobject -property @{
                            Object             = $ADHashDN["$ADUser"].DisplayName
                            UserPrincipalName  = $ADHashDN["$ADUser"].UserPrincipalName
                            PrimarySMTPAddress = $ADHashDN["$ADUser"].PrimarySMTPAddress
                            Granted            = $UserGroupHash[$Member].DisplayName
                            GrantedUPN         = $UserGroupHash[$Member].UserPrincipalName
                            GrantedSMTP        = $UserGroupHash[$Member].PrimarySMTPAddress
                            Checking           = $HasPerm
                            TypeDetails        = "GroupMember"
                            DisplayType        = $ADHashDisplay."$($ADHash["$HasPerm"].msExchRecipientDisplayType)"
                            Permission         = "FullAccess"
                        }
                    }
                }
                elseif ( $ADHash["$HasPerm"].objectClass -notmatch 'group') {
                    New-Object -TypeName psobject -property @{
                        Object             = $ADHashDN["$ADUser"].DisplayName
                        UserPrincipalName  = $ADHashDN["$ADUser"].UserPrincipalName
                        PrimarySMTPAddress = $ADHashDN["$ADUser"].PrimarySMTPAddress
                        Granted            = $ADHash["$HasPerm"].DisplayName
                        GrantedUPN         = $ADHash["$HasPerm"].UserPrincipalName
                        GrantedSMTP        = $ADHash["$HasPerm"].PrimarySMTPAddress
                        Checking           = $_.User
                        TypeDetails        = $ADHashType."$($ADHash["$HasPerm"].msExchRecipientTypeDetails)"
                        DisplayType        = $ADHashDisplay."$($ADHash["$HasPerm"].msExchRecipientDisplayType)"
                        Permission         = "FullAccess"
                    }
                }
            }
        }
    }
    end {

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
        $ServerName = Read-Host "Type then name of the Exchange Server and hit enter"
        Connect-Exchange -Server $ServerName
    }

}
Get-Answer
Get-MailboxMoveOnPremisesPermissionReport -ReportPath ([Environment]::GetFolderPath("Desktop")) -Verbose





#######


