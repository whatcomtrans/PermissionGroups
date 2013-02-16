<#
.SYNOPSIS
TODO

.EXAMPLE
TODO
#>

<#
.SYNOPSIS
Creates a new shared mailbox and an associated FullAccess group to manage access.

.EXAMPLE
TODO
#>
function New-SharedMailbox {
	[CmdletBinding(SupportsShouldProcess=$true)]
	Param(
		[Parameter(Mandatory=$true,Position=0,HelpMessage="Mailbox alias")] 
            [String]$Alias,
		[Parameter(Mandatory=$true,Position=1,HelpMessage="Mailbox name")] 
            [String]$Name,
        [Parameter(Mandatory=$true,Position=2,HelpMessage="Email address")] 
            [String]$EmailAddress,
		[Parameter(Mandatory=$true,Position=3,HelpMessage="Turn on(true)/off(false) automapping, defaults to True")] 
            [Switch]$AutoMapping = $true,
        [Parameter(Mandatory=$false,Position=4,HelpMessage="The OU where the permissions groups will be created")] 
            [String]$PermissionsOU = "",
        [Parameter(Mandatory=$false,Position=5,HelpMessage="If using DirSync, specify the computername where it runs")] 
            [String]$DirSyncHost = ""
	)
	Process {
        [String] $_mailboxalias = $Alias
        [String] $_mailboxname = $Name
        [String] $_PermissionGroup = "SHMB-$_mailboxalias"
        [boolean] $_AutoMapping = $AutoMapping
        [String] $_AutoMapIdicator = ""
        [String] $EmailDomain = ($EmailAddress.Split('@'))[1]

        if ($_AutoMapping) {
            $_AutoMapIdicator = "Users will be recursively AutoMapped to the account."
        }

        #Create security group
        $_GroupDescription = "Users and groups have FullAccess and SendAs permissions to the shared mailbox: $_mailboxalias .  $_AutoMapIdicator"
        [Microsoft.ActiveDirectory.Management.ADGroup]$_newGroup = New-ADGroup -Path $PermissionsOU -Name $_PermissionGroup -SamAccountName $_PermissionGroup -Description $_GroupDescription -GroupScope Global -PassThru

        #Allow time for the change to sync in
        Sleep 30

        #Change security group to distribution group
        Enable-SecurityGroupAsDistributionGroup -Identity $_PermissionGroup -DisplayName $_PermissionGroup -EmailAddress "$_PermissionGroup@$EmailDomain" -Hide

        #Allow time for the change to sync in
        Sleep 30

        #Force directory sync.  Only run if DirSyncHost defined.  This allows this module to work with or without Office365 dirsync.
        if ($DirSyncHost.lenght -ge 1) {
            Force-DirSync -ComputerName $DirSyncHost
            #Allow time for the change to sync in
            Sleep 30
        }

        #Create a shared mailbox
        New-Mailbox -Name $_mailboxname -Alias $_mailboxalias -Shared
        Add-ProxyAddress $_mailboxalias -ProxyAddress "$EmailAddress" -IsDefault

        #TODO: Don't need the first line as these are defaults BUT, should I specify non defaults and should I do that here.
        #      As for the second line, is that even needed or will Office365 setup defaults for me too?  TODO - test
        #Set-Mailbox $_mailboxalias -RoleAssignmentPolicy "WTA Users" -RetentionPolicy "WTA Primary" #-EmailAddresses "$_mailboxalias@ridewta.com", "$_mailboxalias@whatcomtrans.net", ((Get-Mailbox $_mailboxalias).EmailAddresses).toLower()
        #Set-Mailbox $_mailboxalias -ProhibitSendReceiveQuota 5GB -ProhibitSendQuota 4.75GB -IssueWarningQuota 4.5GB

        #Assign the security group the fullAccess permission to access the shared mailbox
        Add-MailboxPermission -Identity $_mailboxalias -User $_PermissionGroup -AccessRights FullAccess -AutoMapping:$_AutoMapping

        #Assign the security gorup the SendAs permission to the shared mailbox
        Add-RecipientPermission -Identity $_mailboxalias -Trustee $_PermissionGroup -AccessRights SendAs -Confirm:$false

        return $_newGroup
        #Done
	}
}

<#
.SYNOPSIS
Syncronizes a shared (or really any) mailbox's permissions so that members of groups assigned to the mailbox when are setup for AutoMapping are added directly to the permissions list with FullAccess.

.EXAMPLE
Sync-SharedMailboxAutoMapping IT@contosu.com
#>
function Sync-SharedMailboxAutoMapping {
	[CmdletBinding(SupportsShouldProcess=$true)]
	Param(
		[Parameter(Mandatory=$true,Position=0,HelpMessage="Mailbox identity")] 
            [Object]$Identity
	)
	Process {
        #TODO - Modify comparison process to handle access other then FullAccess
        #TODO - Also, this assumes the only users directly mapped will be AutoMapped
        #TODO - Add support for confirm and whatif
        #TODO - I am still not confident I have tested every scenario and until I get the above TODO implemented, all changes REQUIRE confirmation
        
        [String[]]$_PermissionGroupUsers = $null

        #Step 1:  Get the SHMB groups associated with the mailbox
        [String[]] $_SHMBGroupNames = (Get-MailboxPermission $Identity | Where-Object -Property IsInherited -EQ -Value $false).User
        if ($_SHMBGroupNames.Count -gt 0) {
            [String[]] $_SHMBGroups = $null
            $_SHMBGroupNames | ForEach-Object {$_SHMBGroups += ((Get-DistributionGroup -Identity $_ -ErrorAction SilentlyContinue).Name)}

            foreach ($_SHMBGroup in $_SHMBGroups) {
                #Step 2:  Verify the group is set to AutoMapping by checking description by -like "* AutoMapped *"
                if ((Get-ADGroup $_SHMBGroup -Properties Description).Description -like "* AutoMapped *") {
                    #Step 3:  Get the users in the group
                    $_PermissionGroupUsers += ((Get-ADGroupMember -Identity $_SHMBGroup -Recursive) | Where-Object -Property objectClass -EQ -Value user).SamAccountName
                }
            }

            #Step 4:  Get the users already assigned permissions
            $_ExistingUsers = ((Get-MailboxPermission $Identity).User | Get-Mailbox -ErrorAction SilentlyContinue).Alias
            
            if ($_ExistingUsers.Count -gt 0 -and $_PermissionGroupUsers.Count -gt 0) {

                #Step 5:  Compare Lists
                $_Comparison = Compare-Object -ReferenceObject $_PermissionGroupUsers -DifferenceObject $_ExistingUsers -IncludeEqual

                #Step 6:  Remove users
                if (($_Comparison | Where-Object -Property SideIndicator -EQ -Value "=>" | Measure-Object).Count -gt 0) {
                    ($_Comparison | Where-Object -Property SideIndicator -EQ -Value "=>").InputObject | ForEach-Object {$_UserMBPermissions = Get-MailboxPermission -Identity $Identity -User $_; $_UserMBPermissions | ForEach-Object {Remove-MailboxPermission -Identity $_.Identity -User $_.User -AccessRights $_.AccessRights -Confirm:$true}}
                }

                #Step 7:  Add users
                if (($_Comparison | Where-Object -Property SideIndicator -EQ -Value "<=" | Measure-Object).Count -gt 0) {
                    ($_Comparison | Where-Object -Property SideIndicator -EQ -Value "<=").InputObject | ForEach-Object {Add-MailboxPermission -Identity $Identity -User $_ -AccessRights FullAccess -AutoMapping:$true -Confirm:$true}
                }
            } else {
                #Step 6:  Remove users

                if ($_ExistingUsers.Count -gt 0) {
                    $_ExistingUsers | ForEach-Object {$_UserMBPermissions = Get-MailboxPermission -Identity $Identity -User $_; $_UserMBPermissions | ForEach-Object {Remove-MailboxPermission -Identity $_.Identity -User $_.User -AccessRights $_.AccessRights -Confirm:$true}}
                }

                #Step 7:  Add users
                if ($_PermissionGroupUsers.Count -gt 0) {
                    $_PermissionGroupUsers | ForEach-Object {Add-MailboxPermission -Identity $Identity -User $_ -AccessRights FullAccess -AutoMapping:$true -Confirm:$true}
                }
            }

            return $true
        } else {
            return $false
        }
	}
}
#>

<#
.SYNOPSIS
Creates a new shared mailbox and an associated FullAccess group to manage access.

.EXAMPLE
TODO
#>
function Add-SharedMailboxGroup {
	[CmdletBinding(SupportsShouldProcess=$true)]
	Param(
		[Parameter(Mandatory=$true,Position=0,HelpMessage="Mailbox alias")] 
            [String]$Alias,
		[Parameter(Mandatory=$true,Position=1,HelpMessage="Mailbox name")] 
            [String]$Name,
        [Parameter(Mandatory=$true,Position=2,HelpMessage="Email address")] 
            [String]$EmailAddress,
		[Parameter(Mandatory=$true,Position=3,HelpMessage="Turn on(true)/off(false) automapping, defaults to True")] 
            [Switch]$AutoMapping = $true,
        [Parameter(Mandatory=$false,Position=4,HelpMessage="The OU where the permissions groups will be created")] 
            [String]$PermissionsOU = "",
        [Parameter(Mandatory=$false,Position=5,HelpMessage="If using DirSync, specify the computername where it runs")] 
            [String]$DirSyncHost = ""
	)
	Process {
        [String] $_mailboxalias = $Alias
        [String] $_mailboxname = $Name
        [String] $_PermissionGroup = "SHMB-$_mailboxalias"
        [boolean] $_AutoMapping = $AutoMapping
        [String] $_AutoMapIdicator = ""
        [String] $EmailDomain = ($EmailAddress.Split('@'))[1]

        if ($_AutoMapping) {
            $_AutoMapIdicator = "Users will be recursively AutoMapped to the account."
        }

        #Create security group
        $_GroupDescription = "Users and groups have FullAccess and SendAs permissions to the shared mailbox: $_mailboxalias .  $_AutoMapIdicator"
        [Microsoft.ActiveDirectory.Management.ADGroup]$_newGroup = New-ADGroup -Path $PermissionsOU -Name $_PermissionGroup -SamAccountName $_PermissionGroup -Description $_GroupDescription -GroupScope Global -PassThru

        #Allow time for the change to sync in
        Sleep 30

        #Change security group to distribution group
        Enable-SecurityGroupAsDistributionGroup -Identity $_PermissionGroup -DisplayName $_PermissionGroup -EmailAddress "$_PermissionGroup@$EmailDomain" -Hide

        #Allow time for the change to sync in
        Sleep 30

        #Force directory sync.  Only run if DirSyncHost defined.  This allows this module to work with or without Office365 dirsync.
        if ($DirSyncHost.lenght -ge 1) {
            Force-DirSync -ComputerName $DirSyncHost
            #Allow time for the change to sync in
            Sleep 30
        }

        #Create a shared mailbox
        New-Mailbox -Name $_mailboxname -Alias $_mailboxalias -Shared
        Add-ProxyAddress $_mailboxalias -ProxyAddress "$EmailAddress" -IsDefault

        #TODO: Don't need the first line as these are defaults BUT, should I specify non defaults and should I do that here.
        #      As for the second line, is that even needed or will Office365 setup defaults for me too?  TODO - test
        #Set-Mailbox $_mailboxalias -RoleAssignmentPolicy "WTA Users" -RetentionPolicy "WTA Primary" #-EmailAddresses "$_mailboxalias@ridewta.com", "$_mailboxalias@whatcomtrans.net", ((Get-Mailbox $_mailboxalias).EmailAddresses).toLower()
        #Set-Mailbox $_mailboxalias -ProhibitSendReceiveQuota 5GB -ProhibitSendQuota 4.75GB -IssueWarningQuota 4.5GB

        #Assign the security group the fullAccess permission to access the shared mailbox
        Add-MailboxPermission -Identity $_mailboxalias -User $_PermissionGroup -AccessRights FullAccess -AutoMapping:$_AutoMapping

        #Assign the security gorup the SendAs permission to the shared mailbox
        Add-RecipientPermission -Identity $_mailboxalias -Trustee $_PermissionGroup -AccessRights SendAs -Confirm:$false

        return $_newGroup
        #Done
	}
}


<#
.SYNOPSIS
TODO

.EXAMPLE
TODO#>
<#
function Verb-Noun {
	[CmdletBinding(SupportsShouldProcess=$true)]
	Param(
		[Parameter(Mandatory=$true,Position=0,HelpMessage="Permission (RW,RO,LO, or a FileSystemRights Enumeration value, see http://msdn.microsoft.com/en-us/library/system.security.accesscontrol.filesystemrights.aspx)")] 
            [String]$Permission,
		[Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=1,HelpMessage="Path to set permission on.")] 
            [String]$Path,
		[Parameter(Mandatory=$true,Position=2,HelpMessage="Name of user or group to assign permission to.")] 
            [String]$AssignedTo,
		[Parameter(Mandatory=$false,Position=3,HelpMessage="Determine how this rule is inherited by child objects.  Values are None, ContainerInherit, ObjectInherit or some combination of these values in a comma seperated string.  Default is ContainerInherit, ObjectInherit.  See http://msdn.microsoft.com/en-us/magazine/cc163885.aspx#S3")] 
            [String]$InheritanceFlags="ContainerInherit, ObjectInherit",
		[Parameter(Mandatory=$false,Position=4,HelpMessage="Determine how inheritance of this rule is propagated to child objects.  Values are None, NoPropagateInherit and InheritOnly or some combination of these values in a comma seperated string.  Default is None.  See http://msdn.microsoft.com/en-us/magazine/cc163885.aspx#S3")] 
            [String]$PropagationFlags="None",
		[Parameter(Mandatory=$false,Position=5,HelpMessage="Whether to Allow or Deny the permission, defaults to Allow")] [ValidateSet("Allow","Deny")] 
            [String]$Grant="Allow"
	)
	Begin {
		#TODO
	}
	Process {
        #TODO
	}
}
#>

Export-ModuleMember -Function "New-SharedMailbox","Sync-SharedMailboxAutoMapping" #TODO "Verb-Noun"