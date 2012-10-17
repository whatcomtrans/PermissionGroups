<#
.SYNOPSIS
TODO

.EXAMPLE
TODO
#>

#Shared variables
$DirSyncHost = "SRVMSOL1"
$PermissionsOU = "OU=PermissionGroups,DC=whatcomtrans,DC=net"

<#
.SYNOPSIS
TODO

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
		[Parameter(Mandatory=$true,Position=2,HelpMessage="Turn on(true)/off(false) automapping, defaults to True")] 
            [Switch]$AutoMapping = $true
	)
	Process {
        [String] $_mailboxalias = $Alias
        [String] $_mailboxname = $Name
        [String] $_PermissionGroup = "SHMB-$_mailboxalias"
        [boolean] $_AutoMapping = $AutoMapping
        [String] $_AutoMapIdicator = ""

        if ($_AutoMapping) {
            $_AutoMapIdicator = "Users will be recursively AutoMapped to the account."
        }

        #Create security group
        [Microsoft.ActiveDirectory.Management.ADGroup]$_newGroup = New-ADGroup -Path $PermissionsOU -Name $_PermissionGroup -SamAccountName $_PermissionGroup -Description "Users and groups have FullAccess and SendAs permissions to the shared mailbox: $_mailboxalais.  $_AutoMapIndicator" -GroupScope Universal -PassThru

        #Allow time for the change to sync in
        Sleep 30

        #TODO - Make this more generic
        Enable-SecurityGroupAsDistributionGroup -Identity $_PermissionGroup -DisplayName $_PermissionGroup -EmailAddress "$_PermissionGroup@whatcomtran.net" -Hide

        #Allow time for the change to sync in
        Sleep 30

        #Force directory sync
        #TODO - Make this more generic
        $scb = {
            #Force DirSync
            Add-PSSnapin Coexistence-Configuration
            Start-OnlineCoexistenceSync
            Sleep 30
        }
        Invoke-Command -ComputerName $DirSyncHost -ScriptBlock $scb

        #Allow time for the change to sync in
        Sleep 30

        #Create a shared mailbox
        New-Mailbox -Name $_mailboxname -Alias $_mailboxalias -Shared
        Set-Mailbox $_mailboxalias -ProhibitSendReceiveQuota 5GB -ProhibitSendQuota 4.75GB -IssueWarningQuota 4.5GB

        #TODO - Make this more generic
        Add-ProxyAddress $_mailboxalias -ProxyAddress "$_mailboxalias@ridewta.com"
        Add-ProxyAddress $_mailboxalias -ProxyAddress "$_mailboxalias@whatcomtrans.net" -IsDefault
        Set-Mailbox $_mailboxalias -RoleAssignmentPolicy "WTA Users" -RetentionPolicy "WTA Primary" #-EmailAddresses "$_mailboxalias@ridewta.com", "$_mailboxalias@whatcomtrans.net", ((Get-Mailbox $_mailboxalias).EmailAddresses).toLower()

        #Assign the security group the fullAccess permission to access the shared mailbox
        Add-MailboxPermission -Identity $_mailboxalias -User $_PermissionGroup -AccessRights FullAccess -AutoMapping:$_AutoMapping

        #Assign the security gorup the SendAs permission to the shared mailbox
        Add-RecipientPermission -Identity $_mailboxalias -Trustee $_PermissionGroup -AccessRights SendAs

        if ($_AutoMapping) {
            Sync-SharedMailboxAutoMapping $_mailboxalias
        }

        return $_newGroup
        #Done
	}
}

<#
.SYNOPSIS
TODO

.EXAMPLE
TODO
#>
function Sync-SharedMailboxAutoMapping {
	[CmdletBinding(SupportsShouldProcess=$true)]
	Param(
		[Parameter(Mandatory=$true,Position=0,HelpMessage="Mailbox identity")] 
            [String]$Identity<#,
		[Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=1,HelpMessage="Path to set permission on.")] 
            [String]$Path,
		[Parameter(Mandatory=$true,Position=2,HelpMessage="Name of user or group to assign permission to.")] 
            [String]$AssignedTo,
		[Parameter(Mandatory=$false,Position=3,HelpMessage="Determine how this rule is inherited by child objects.  Values are None, ContainerInherit, ObjectInherit or some combination of these values in a comma seperated string.  Default is ContainerInherit, ObjectInherit.  See http://msdn.microsoft.com/en-us/magazine/cc163885.aspx#S3")] 
            [String]$InheritanceFlags="ContainerInherit, ObjectInherit",
		[Parameter(Mandatory=$false,Position=4,HelpMessage="Determine how inheritance of this rule is propagated to child objects.  Values are None, NoPropagateInherit and InheritOnly or some combination of these values in a comma seperated string.  Default is None.  See http://msdn.microsoft.com/en-us/magazine/cc163885.aspx#S3")] 
            [String]$PropagationFlags="None",
		[Parameter(Mandatory=$false,Position=5,HelpMessage="Whether to Allow or Deny the permission, defaults to Allow")] [ValidateSet("Allow","Deny")] 
            [String]$Grant="Allow"#>
	)
	Begin {
		#TODO
	}
	Process {
        #TODO
	}
}
#>


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

Export-ModuleMember -Function "New-SharedMailbox" #TODO "Verb-Noun"