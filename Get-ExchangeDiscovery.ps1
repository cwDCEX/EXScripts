# The following captures Exchange related information from the environment using a combination of AD and Exchange Powershell Commands. 
# Results are saved to a text file in the following location: C:\temp\ExchangeServerInfo.txt.
# "Get-" cmdlets are used to gather information. See inline comments to identify the information that is being gathered.
# This script may take a long time to run on larger environments.

# Basic Exchange Server Info
$Destination = "C:\temp\ExchangeServerInfo.txt"
$Date = Get-Date
$Line = "Exchange Server Information - $($Date)" | Out-File $Destination

$ExchangeServers = Get-ExchangeServer
$LocalExchangeServer = $Env:ComputerName
$ServerVersionCheck = $True

Foreach ($ExchangeServer in $ExchangeServers) {
    $Server = $ExchangeServer.Name
    $Site = $ExchangeServer.Site
    $ServerRole = $ExchangeServer.ServerRole
    $Edition = $ExchangeServer.Edition
    If ($Server -ne $LocalExchangeServer) {
        $Version = Invoke-Command -ComputerName $Server -ScriptBlock {
            $Ver = Get-Command Exsetup.exe | ForEach-Object {$_.FileversionInfo}
            $Version = $Ver.FileVersion
            $Version
        } -ErrorAction Stop
    } Else {
        $Ver = Get-Command Exsetup.exe | ForEach-Object {$_.FileversionInfo}
        $Version = $Ver.FileVersion
    }
    $Line = "$Server|$Version|$Edition|$ServerRole|$Site" | Out-file $Destination -Append
}

# Active Directory and Exchange Schema Information
$Line = "----------------------------------------------" | Out-File $Destination -Append
$Line = "Active Directory and Exchange Schema Information" | Out-File $Destination -Append

# AD FSMO roles
$Line = "AD FSMO Roles - DNS Root:" | Out-File $Destination -Append
$Line = (Get-ADForest).RootDomain | Out-File $Destination -Append

$Line = "AD FSMO Roles - DomainNamingMaster:" | Out-File $Destination -Append
$Line = (Get-ADForest).DomainNamingMaster | Out-File $Destination -Append

$Line = "AD FSMO Roles - InfrastructureMaster:" | Out-File $Destination -Append
$Line = (Get-ADDomain).InfrastructureMaster | Out-File $Destination -Append

$Line = "AD FSMO Roles - PDCEmulator:" | Out-File $Destination -Append
$Line = (Get-ADDomain).PDCEmulator | Out-File $Destination -Append

$Line = "AD FSMO Roles - RIDMaster:" | Out-File $Destination -Append
$Line = (Get-ADDomain).RIDMaster | Out-File $Destination -Append

$Line = "AD FSMO - SchemaMaster:" | Out-File $Destination -Append
$Line = (Get-ADForest).SchemaMaster | Out-File $Destination -Append

# Exchange Schema Version
$Line = "Exchange Schema Version" | Out-File $Destination -Append
$sc = (Get-ADRootDSE).SchemaNamingContext
$ob = "CN=ms-Exch-Schema-Version-Pt," + $sc
$Line = ((Get-ADObject $ob -pr rangeUpper).rangeUpper) | Out-File $Destination -Append

# Exchange Object Version (domain)
$Line = "Exchange Object Version (domain):" | Out-File $Destination -Append
$dc = (Get-ADRootDSE).DefaultNamingContext
$ob = "CN=Microsoft Exchange System Objects," + $dc
$line = ((Get-ADObject $ob -pr objectVersion).objectVersion) | Out-File $Destination -Append

# Exchange Object Version (forest)
$Line = "Exchange Object Version (forest):" | Out-File $Destination -Append
$cc = (Get-ADRootDSE).ConfigurationNamingContext
$fl = "(objectClass=msExchOrganizationContainer)"
$Line = ((Get-ADObject -LDAPFilter $fl -SearchBase $cc -pr objectVersion).objectVersion) | Out-File $Destination -Append

# AD Forest functional level
$Line = "AD Forest Functional Level:" | Out-File $Destination -Append
$Line = (Get-ADForest).ForestMode | Out-File $Destination -Append

# Current Domain Controllers and their versions
$Line = "Current Domain Controllers and their versions:" | Out-File $Destination -Append
$DomainControllers = Get-ADDomainController -Filter * | Select-Object Name, OperatingSystem
$Line = $DomainControllers | Format-Table -AutoSize -Wrap | Out-File $Destination -Append

# AD IP Site links
$Line = "AD IP Site Links:" | Out-File $Destination -Append
$ADSiteLinks = Get-ADReplicationSiteLink -Filter * | Select-Object Name, Cost, ReplicationFrequencyInMinutes, SitesIncluded
$Line = $ADSiteLinks | Format-Table -AutoSize -Wrap | Out-File $Destination -Append

# Windows Server Configuration, Hardware/Resources, Pagefile, Event Logs, OS Version
$Line = "----------------------------------------------" | Out-File $Destination -Append
$Line = "Windows Server Configuration, Hardware/Resources, Pagefile, Event Logs, OS Version" | Out-File $Destination -Append
# Processor Cores - Logical and Physical
$Processors = Get-WMIObject Win32_Processor -ComputerName $Server
$LogicalCPU = ($Processors | Measure-Object -Property NumberOfLogicalProcessors -sum).Sum
$PhysicalCPU = ($Processors | Measure-Object -Property NumberOfCores -sum).Sum

# Server RAM
$RamInGb = (Get-wmiobject -ComputerName $Server -Classname win32_physicalmemory -ErrorAction Stop | measure-object -property capacity -sum).sum/1GB

# Pagefile Configuration
$PageFileCheck = Get-CIMInstance -ComputerName $Server -Class WIN32_PageFile -ErrorAction STOP
$Managed = $False
If ($Null -ne $PageFileCheck) {
$MaximumSize = (Get-CimInstance -ComputerName $Server -Query "Select * from win32_PageFileSetting" | select-object MaximumSize).MaximumSize
$InitialSize = (Get-CimInstance -ComputerName $Server -Query "Select * from win32_PageFileSetting" | select-object InitialSize).InitialSize
}

# OS Version
$OSVersion = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $Server | Select-Object Caption, Version

# Create an output object for the information
$ServerOutput = [PSCustomObject]@{
    ServerName = $Server
    LogicalCores = $LogicalCPU
    PhysicalCores = $PhysicalCPU
    RAM = $RamInGb
    OSVersion = $OSVersion.Caption
    OSVersionNumber = $OSVersion.Version
    PagefileManaged = $Managed
    PagefileInitialSize = $InitialSize
    PagefileMaximumSize = $MaximumSize
}

# Append the server information to the text file
$ServerOutput | Format-Table -AutoSize -Wrap | Out-File -FilePath "C:\temp\ExchangeServerInfo.txt" -Append

# Exchange Server Roles and Service Statues
$Line = "----------------------------------------------" | Out-File $Destination -Append
$Line = "Server Roles" | Out-File $Destination -Append
$Line = Get-ExchangeServer | ft Name, Domain, Edition, AdminDisplayVersion, ServerRole -AutoSize -Wrap | Out-File $Destination -Append
$Line = "----------------------------------------------" | Out-File $Destination -Append
$Line = "Services Status" | Out-File $Destination -Append
$Line = Get-Service MSExchange* | ft Name, DisplayName, Status, StartType -AutoSize -Wrap | Out-File $Destination -Append

# Exchange Domains
$Line = "----------------------------------------------" | Out-File $Destination -Append
$Line = "Exchange Accepted Domains and Remote Domains" | Out-File $Destination -Append

# Exchange Accepted Domains
$Line = "Exchange Accepted Domains:" | Out-File $Destination -Append
$AcceptedDomains = Get-AcceptedDomain | Select-Object Name, DomainName, DomainType, Default
$Line = $AcceptedDomains | Format-Table -AutoSize | Out-File $Destination -Append

# Exchange Remote Domains
$Line = "Exchange Remote Domains:" | Out-File $Destination -Append
$RemoteDomains = Get-RemoteDomain | Select-Object Name, DomainName, IsInternal, IsInboundEnabled
$Line = $RemoteDomains | Format-Table -AutoSize | Out-File $Destination -Append

# Exchange Certificates
$Line = "----------------------------------------------" | Out-File $Destination -Append
$Line = "Exchange Certificates" | Out-File $Destination -Append
$ExchangeCertificates = Get-ExchangeCertificate | Select-Object Thumbprint, Services, Issuer, Subject, NotAfter, Status
$Line = $ExchangeCertificates | Format-Table -AutoSize | Out-File $Destination -Append

# Output URLs for various virtual directories and client access services
$Line = "----------------------------------------------" | Out-File $Destination -Append
$Line = "Virtual Directories/CAS URLs" | Out-File $Destination -Append
$Destination = "ExchangeServerInfo.txt"

$Line = "OWA Virtual Directories" | Out-File $Destination -Append
Get-OWAVirtualDirectory -ADPropertiesOnly | ft Server, *lurl* -Auto | Out-File $Destination -Append

$Line = "EWS Virtual Directories" | Out-File $Destination -Append
Get-WebServicesVirtualDirectory -ADPropertiesOnly | ft Server, *lurl* -Auto | Out-File $Destination -Append

$Line = "ActiveSync Virtual Directories" | Out-File $Destination -Append
Get-ActiveSyncVirtualDirectory -ADPropertiesOnly | ft Server, *lurl* -Auto | Out-File $Destination -Append

$Line = "AutoDiscover Virtual Directories" | Out-File $Destination -Append
Get-AutoDiscoverVirtualDirectory -ADPropertiesOnly | ft Server, *lurl* -Auto | Out-File $Destination -Append

$Line = "MAPI Virtual Directories" | Out-File $Destination -Append
Get-MAPIVirtualDirectory -ADPropertiesOnly | ft Server, *lurl* -Auto | Out-File $Destination -Append

$Line = "Offline Address Book Virtual Directories" | Out-File $Destination -Append
Get-OABVirtualDirectory -ADPropertiesOnly | ft Server, *lurl* -Auto | Out-File $Destination -Append

$Line = "Autodiscover Internal Uri" | Out-File $Destination -Append
Get-ClientAccessService | ft Name, *uri* -Auto | Out-File $Destination -Append

$Line = "All Client Access Settings" | Out-File $Destination -Append
Get-ClientAccessService | SELECT * | Format-List | Out-File $Destination -Append

# Email Address Policies
$Line = "----------------------------------------------" | Out-File $Destination -Append
$Line = "Email Address Policies" | Out-File $Destination -Append
$Line = Get-EmailAddressPolicy -ErrorAction STOP | Ft Name, RecipientFilterType, Priority, Enabled*, RecipientFilter, LDAP*, IsValid -AutoSize -Wrap | Out-File $Destination -Append

# Address Book Policies
$Line = "----------------------------------------------" | Out-File $Destination -Append
$Line = "Address Book Policies" | Out-File $Destination -Append
$Line = Get-AddressBookPolicy | FT Identity,Bindings,EnabledAuthMechanism,MaxMessageSize,PermissionGroups -AutoSize -Wrap | Out-File $Destination -Append

# Receive Connectors
$Line = "----------------------------------------------" | Out-File $Destination -Append
$Line = "Receive Connectors" | Out-File $Destination -Append
$ReceiveConnectors = Get-ReceiveConnector | Select-Object Identity, Bindings, @{Name='RemoteIPRanges'; Expression={($_.RemoteIPRanges -join ", ")}}, EnabledAuthMechanism, MaxMessageSize, PermissionGroups, RequireTLS, TransportRole
$Line = $ReceiveConnectors | Format-Table -AutoSize -Wrap | Out-file $Destination -Append

# Send Connectors
$Line = "----------------------------------------------" | Out-File $Destination -Append
$Line = "Send Connectors" | Out-File $Destination -Append
$Line = Get-SendConnector -ErrorAction STOP | Ft Identity, HomeMtaServerId, Enabled, MaxMessageSize, AddressSpaces, CloudServicesMailEnabled, RequireTLS, SmartHosts -AutoSize -Wrap | Out-file $Destination -Append

# DLP Policies
$Line = "----------------------------------------------" | Out-File $Destination -Append
$Line = "DLP Policies" | Out-File $Destination -Append
$Line = Get-DLPPolicy | Ft Name, State, Mode, Identity -AutoSize -Wrap | Out-File $Destination -Append

# Retention Settings
$Line = "----------------------------------------------" | Out-File $Destination -Append
$Line = "Retention Settings" | Out-File $Destination -Append
$Line = Get-OrganizationConfig | Select-Object DefaultPublicFolderProhibitPostQuota, DefaultPublicFolderIssueWarningQuota, DefaultPublicFolderMaxItemSize, DefaultPublicFolderDeletedItemRetention | Fl | Out-File $Destination -Append

$Line = "----------------------------------------------" | Out-File $Destination -Append
$Line = "Quota Limits per Database" | Out-File $Destination -Append
$Line = Get-MailboxDatabase | Ft Name, IssueWarningQuota, ProhibitSendQuota, ProhibitSendReceiveQuota, RecoverableItemsQuota, RecoverableItemsWarningQuota -AutoSize -Wrap | Out-File $Destination -Append

# Exchange Mailbox Databases
$Line = "----------------------------------------------" | Out-File $Destination -Append
$Line = "Exchange Mailbox Databases" | Out-File $Destination -Append
$Line = Get-MailboxDatabase | Ft Name, Server, Recovery, ReplicationType, LogFolderPath, EdbFilePath, DeletedItemRetention, MailboxRetention -AutoSize -Wrap | Out-File $Destination -Append

#Exchange Mailboxes per Database
$Line = "----------------------------------------------" | Out-File $Destination -Append
$Line = "Mailbox per Database" | Out-File $Destination -Append
$MailboxDatabases = Get-MailboxDatabase
foreach ($Database in $MailboxDatabases) {
    $Line = "Database: $($Database.Name)" | Out-File $Destination -Append
    $Line = Get-Mailbox -Database $Database.Name | FT Name, PrimarySMTPAddress, ServerName, ProhibitSendQuota -AutoSize -Wrap | Out-File $Destination -Append
}

#SCP Records
$Line = "----------------------------------------------" | Out-File $Destination -Append
$Line = "SCP Records" | Out-File $Destination -Append
$Line = Get-ClientAccessServer | ft Name, AutoDiscoverServiceInternalUri -AutoSize -Wrap | Out-File $Destination -Append

# Transport Config
$Line = "----------------------------------------------" | Out-File $Destination -Append
$Line = "Transport Configuration" | Out-File $Destination -Append
$Line = Get-TransportConfig | Fl | Out-File $Destination -Append

# Mail Transport Rules
$Line = "----------------------------------------------" | Out-File $Destination -Append
$Line = "Mail Transport Rules" | Out-File $Destination -Append
$Line = Get-TransportRule | FT Name, Priority, Enabled, Description -AutoSize -Wrap | Out-File $Destination -Append

# Exchange RBAC
$Line = "----------------------------------------------" | Out-File $Destination -Append
$Line = "Exchange RBAC Admin roles and assignments" | Out-File $Destination -Append
$Line = Get-RoleGroup | Ft Name, ManagedBy -AutoSize -Wrap | Out-File $Destination -Append
$Line = Get-ManagementRoleAssignment -RoleAssigneeType User -Delegating $False | Select-Object Name, Role, RoleAssigneeName | Format-List | Out-File $Destination -Append

# ActiveSync Policies
$Line = "----------------------------------------------" | Out-File $Destination -Append
$Line = "ActiveSync Device Policies" | Out-File $Destination -Append
$Line = Get-ActiveSyncMailboxPolicy | Ft Name, AllowNonProvisionableDevices, DeviceEncryptionEnabled, RequireDeviceEncryption, PasswordEnabled -AutoSize -Wrap | Out-File $Destination -Append

# OWA Mailbox Policies
$Line = "----------------------------------------------" | Out-File $Destination -Append
$Line = "OWA Mailbox Policies" | Out-File $Destination -Append
$Line = Get-OWAMailboxPolicy | Ft Name, DirectFileAccessOnPublicComputersEnabled, DirectFileAccessOnPrivateComputersEnabled, ForceSaveAttachmentFilteringEnabled, AllowedFileTypes, BlockedFileTypes -AutoSize -Wrap | Out-File $Destination -Append

# Retention Tags and Retention Policies
$Line = "----------------------------------------------" | Out-File $Destination -Append
$Line = "Retention Tags and Retention Policies" | Out-File $Destination -Append
$Line = Get-RetentionPolicyTag | Ft Name, Type, AgeLimitForRetention, RetentionAction -AutoSize -Wrap | Out-File $Destination -Append
$Line = Get-RetentionPolicy | Ft Name, RetentionPolicyTagLinks -AutoSize -Wrap | Out-File $Destination -Append

# Address Lists
$Line = "----------------------------------------------" | Out-File $Destination -Append
$Line = "Address Lists" | Out-File $Destination -Append
$Line = Get-AddressList | Ft Name, DisplayName, RecipientFilter -AutoSize -Wrap | Out-File $Destination -Append

# Offline Address Book
$Line = "----------------------------------------------" | Out-File $Destination -Append
$Line = "Offline Address Book" | Out-File $Destination -Append
$Line = Get-OfflineAddressBook | Ft Name, Versions, AddressLists, PublicFolderDistributionEnabled, WebDistributionEnabled -AutoSize -Wrap | Out-File $Destination -Append

$Line = "----------------------------------------------" | Out-File $Destination -Append
$Line = "Hybrid Configuration" | Out-File $Destination -Append

# Check if the Hybrid Configuration is available
try {
    $HybridConfiguration = Get-HybridConfiguration -ErrorAction Stop
    $Line = $HybridConfiguration | Format-List | Out-File $Destination -Append
} catch {
    $Line = "Hybrid Configuration not found or not applicable." | Out-File $Destination -Append
}

Write-Host "Script Excecution Complete"