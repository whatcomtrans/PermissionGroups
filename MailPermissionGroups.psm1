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
		[Parameter(Mandatory=$true,Position=1,HelpMessage="Mailbox name")] 
            [String]$Name,
        [Parameter(Mandatory=$true,Position=2,HelpMessage="Email address")] 
            [String]$EmailAddress,
		[Parameter(Mandatory=$false,Position=3,HelpMessage="Turn on(true)/off(false) automapping, defaults to True")] 
            [Switch]$AutoMapping = $true,
        [Parameter(Mandatory=$false, ValueFromPipeline=$true, Position=4,HelpMessage="Optional array of members to add (accepts same objects as Add-ADGroupMember)")] 
            [Object[]] $Members,
        [Parameter(Mandatory=$true,Position=5,HelpMessage="The OU where the permissions groups will be created")] 
            [String]$PermissionsOU = "",
        [Parameter(Mandatory=$false,Position=6,HelpMessage="If using DirSync, specify the computername where it runs")] 
            [String]$DirSyncHost = ""
	)
	Process {
        #Determine alias from email address
        [String]$Alias = $EmailAddress.Split("@")[0]

        #Create a shared mailbox
        New-Mailbox -Name $Name -Alias $Alias -Shared
        Add-ProxyAddress $Alias -ProxyAddress "$EmailAddress" -IsDefault

        #Create and associate group
        if (!$Members) {
            return (Add-SharedMailboxGroup -Identity $Name -Permissions "FullAccess" -AutoMapping:$AutoMapping -PermissionsOU $PermissionsOU -DirSyncHost $DirSyncHost) #TODO - Need to add additional parameters
        } else {
            return (Add-SharedMailboxGroup -Identity $Name -Permissions "FullAccess" -AutoMapping:$AutoMapping -Members $Members -PermissionsOU $PermissionsOU -DirSyncHost $DirSyncHost) #TODO - Need to add additional parameters
        }

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
		[Parameter(Mandatory=$true,Position=0,ValueFromPipelineByPropertyName=$True,HelpMessage="Mailbox identity")] 
            [Object]$Identity
	)
	Process {
        #TODO - Modify comparison process to handle access other then FullAccess
        #TODO - Also, this assumes the only users directly mapped will be AutoMapped
        #TODO - Add support for confirm and whatif
        $doConfirm = $false
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
                    ($_Comparison | Where-Object -Property SideIndicator -EQ -Value "=>").InputObject | ForEach-Object {$_UserMBPermissions = Get-MailboxPermission -Identity $Identity -User $_; $_UserMBPermissions | ForEach-Object {Remove-MailboxPermission -Identity $_.Identity -User $_.User -AccessRights $_.AccessRights -Confirm:$doConfirm}}
                }

                #Step 7:  Add users
                if (($_Comparison | Where-Object -Property SideIndicator -EQ -Value "<=" | Measure-Object).Count -gt 0) {
                    ($_Comparison | Where-Object -Property SideIndicator -EQ -Value "<=").InputObject | ForEach-Object {Add-MailboxPermission -Identity $Identity -User $_ -AccessRights FullAccess -AutoMapping:$true -Confirm:$doConfirm}
                }
            } else {
                #Step 6:  Remove users

                if ($_ExistingUsers.Count -gt 0) {
                    $_ExistingUsers | ForEach-Object {$_UserMBPermissions = Get-MailboxPermission -Identity $Identity -User $_; $_UserMBPermissions | ForEach-Object {Remove-MailboxPermission -Identity $_.Identity -User $_.User -AccessRights $_.AccessRights -Confirm:$doConfirm}}
                }

                #Step 7:  Add users
                if ($_PermissionGroupUsers.Count -gt 0) {
                    $_PermissionGroupUsers | ForEach-Object {Add-MailboxPermission -Identity $Identity -User $_ -AccessRights FullAccess -AutoMapping:$true -Confirm:$doConfirm}
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
		[Parameter(Mandatory=$true,Position=0,ValueFromPipelineByPropertyName=$True,HelpMessage="Mailbox Identity")] 
            [String]$Identity,
        [Parameter(Mandatory=$false,Position=1,HelpMessage="Permissions")] 
            [String]$Permissions = "FullAccess",
		[Parameter(Mandatory=$false,Position=2,HelpMessage="Turn on(true)/off(false) automapping, defaults to True")] 
            [Switch]$AutoMapping = $true,
        [Parameter(Mandatory=$false,Position=3,HelpMessage="Optional array of members to add (accepts same objects as Add-ADGroupMember)")] 
            [Object[]] $Members,
        [Parameter(Mandatory=$true,Position=4,HelpMessage="The OU where the permissions groups will be created")] 
            [String]$PermissionsOU = "",
        [Parameter(Mandatory=$false,Position=5,HelpMessage="If using DirSync, specify the computername where it runs")] 
            [String]$DirSyncHost = ""
	)
	Process {
        #TODO - Add support for confirm and whatif
        $doConfirm = $false
        
        $cleanIdentity = $Identity.Replace(' ', '_')

        [String] $_PermissionGroup = "SHMB-" + $cleanIdentity + "-" + $Permissions
        [boolean] $_AutoMapping = $AutoMapping
        [String] $_AutoMapIdicator = ""

        if ($_AutoMapping) {
            $_AutoMapIdicator = "Users will be recursively AutoMapped to the account."
        }

        #Create security group
        $_GroupDescription = "Users and groups have FullAccess and SendAs permissions to the shared mailbox: $Identity .  $_AutoMapIdicator"
        [Microsoft.ActiveDirectory.Management.ADGroup]$_newGroup = New-ADGroup -Path $PermissionsOU -Name $_PermissionGroup -SamAccountName $_PermissionGroup -Description $_GroupDescription -GroupScope Global -PassThru

        #Allow time for the change to sync in AD
        Sleep 60

        #Force directory sync.  Only run if DirSyncHost defined.  This allows this module to work with or without Office365 dirsync.
        if ($DirSyncHost.Length -ge 1) {
            Force-DirSync -ComputerName $DirSyncHost
            #Allow time for the change to sync in Office365
            Sleep 60
        }

        #Get mailbox so I can get email address for distribution group
        [String] $EmailDomain = (((Get-Mailbox $Identity).PrimarySMTPAddress).Split('@'))[1]

        #Change security group to distribution group
        Enable-SecurityGroupAsDistributionGroup -Identity $_PermissionGroup -DisplayName $_PermissionGroup -EmailAddress "$_PermissionGroup@$EmailDomain" -Hide -Confirm:$doConfirm

        #Allow time for the change to sync in AD
        Sleep 60

        #Force directory sync.  Only run if DirSyncHost defined.  This allows this module to work with or without Office365 dirsync.
        if ($DirSyncHost.Length -ge 1) {
            Force-DirSync -ComputerName $DirSyncHost
            #Allow time for the change to sync in Office365
            Sleep 60
        }

        #Assign the security group the fullAccess permission to access the shared mailbox
        Add-MailboxPermission -Identity $Identity -User $_PermissionGroup -AccessRights $Permissions -AutoMapping:$_AutoMapping -Confirm:$doConfirm

        If ($Permissions -eq "FullAccess") { 
            #Assign the security gorup the SendAs permission to the shared mailbox
            Add-RecipientPermission -Identity $Identity -Trustee $_PermissionGroup -AccessRights SendAs -Confirm:$doConfirm
        }

        if ($Members) {
            #Add members to the group
            $_newGroup | Add-ADGroupMember -Members $Members
            
            #Allow time for the change to sync in AD
            Sleep 30

            #Force directory sync.  Only run if DirSyncHost defined.  This allows this module to work with or without Office365 dirsync.
            if ($DirSyncHost.Length -ge 1) {
                Force-DirSync -ComputerName $DirSyncHost
                #Allow time for the change to sync in Office365
                Sleep 60
            }

            if ($_AutoMapping) {
                Sync-SharedMailboxAutoMapping $Identity
            }
        }

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

Export-ModuleMember -Function "New-SharedMailbox","Sync-SharedMailboxAutoMapping","Add-SharedMailboxGroup" #TODO "Verb-Noun"