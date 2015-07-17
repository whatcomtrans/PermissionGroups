
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

        Sleep 120

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

function New-PermissionsDistributionGroup {
	[CmdletBinding(SupportsShouldProcess=$true)]
	Param(
		[Parameter(Mandatory=$true,Position=1,HelpMessage="Mailbox name")] 
            [String]$Name,
        [Parameter(Mandatory=$true,Position=2,HelpMessage="Email address")] 
            [String]$DisplayName,
        [Parameter(Mandatory=$false, ValueFromPipeline=$true, Position=4,HelpMessage="Optional array of members to add (accepts same objects as Add-ADGroupMember)")] 
            [Object[]] $Members,
        [Parameter(Mandatory=$true,HelpMessage="The OU where the permissions groups will be created")] 
            [String]$OU = "",
        [Parameter(Mandatory=$false,HelpMessage="A prefix that exists on the ADGroup but not on the DistributionGroup name.  Defaults to 'EL-'")] 
            [String]$ADGroupPrefix = "EL-",
        [Parameter(Mandatory=$true,HelpMessage="A prefix that exists on the ADGroup but not on the DistributionGroup name.  Defaults to 'EL-'")] 
            [String]$PrimarySMTPDomain,
        [Parameter(HelpMessage="Returns the new AD Group, defaults to returning the distribution list object.")] 
            [Switch]$ReturnADGroup

	)
	Process {
        #Setup names.
        [String] $_cleanIdentity = $Name.Replace(' ', '_')
        [String] $_PermissionGroup = "$($ADGroupPrefix)$($_cleanIdentity)"
        [String] $_GroupDescription = "This group is synced with an Exchange Distribution list of similar name."

        [String]$Alias = $_cleanIdentity  #$EmailAddress.Split("@")[0]

        if ($pscmdlet.ShouldProcess("Create AD Group $_PermissionGroupName and Exchange Distribution Group $_DistListName adding members to both.")) {
            #Create the Permission Group
            [Microsoft.ActiveDirectory.Management.ADGroup]$PermissionGroup = New-ADGroup -Path $OU -Name $_PermissionGroup -SamAccountName $_PermissionGroup -Description $_GroupDescription -GroupScope Global -PassThru

            #Add Members to the Permission Group
            if ($Members) {
                Add-ADGroupMember -Identity $PermissionGroup.SamAccountName -Members $Members
            }

            #Create the New-DistributionGroup
            $DistGroup = New-DistributionGroup -Name $DisplayName -Alias $Alias -DisplayName $DisplayName -PrimarySmtpAddress "$($Alias)@$($PrimarySMTPDomain)" -MemberDepartRestriction Closed

            #Sync the membership by calling Sync-PermissionsDistributionList
            Sync-PermissionsDistributionGroup -PermissionGroup $PermissionGroup -DistributionGroup $DistGroup

            if ($ReturnADGroup) {
                return $PermissionGroup
            } else {
                return $DistGroup
            }
        }
	}
}

function Sync-PermissionsDistributionGroup {
	[CmdletBinding(SupportsShouldProcess=$true,DefaultParameterSetName="objects")]
	Param(
        [Parameter(Mandatory=$true,Position=1,ValueFromPipeline=$true,ParameterSetName="objects",HelpMessage="ADGroup PermissionGroup object")] 
            [Microsoft.ActiveDirectory.Management.ADGroup[]]$PermissionGroup,
        [Parameter(Mandatory=$true,Position=2,ValueFromPipeline=$true,ParameterSetName="objects",HelpMessage="Exchnage DistributionGroup object")] 
            [PSObject[]]$DistributionGroup,
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ParameterSetName="OU",HelpMessage="The OU where the permissions groups can be found.  All will be synced.")] 
            [String]$OU = "",
        [Parameter(Mandatory=$false,ParameterSetName="OU",HelpMessage="A prefix that exists on the ADGroup but not on the DistributionGroup name.  Defaults to 'EL-'")] 
            [String]$ADGroupPrefix = "EL-",
        [Parameter(Mandatory=$false,ParameterSetName="OU",HelpMessage="The ADGroup property to use to find the email distribution group.  Defaults to Name.")] 
            [String]$ADGroupProperty = "Name",
        [Parameter(Mandatory=$false,HelpMessage="Default syncs ADGroup to DistributionGroup, this reverse the sync direction.")] 
            [Switch] $ReverseDirection,
        [Parameter(Mandatory=$false,HelpMessage="Default will get all ADGroup members recursevly and adds them individually, this will add as is.")] 
            [Switch] $DoNotFlatten
	)
	Process {
        [System.Collections.ArrayList] $_pairs = New-Object System.Collections.ArrayList

        if ($OU) {
            #OU
            $_PermissionGroups = Get-ADGroup -SearchScope OneLevel -SearchBase $OU -Filter "*" | Where-Object -Property Name -Value "$($ADGroupPrefix)*" -like | Get-ADGroup
            forEach ($_ADGroup in $_PermissionGroups) {
                if ($ADGroupPrefix) {
                    $_DGroupName = $_ADGroup.$ADGroupProperty.Replace($ADGroupPrefix, "")
                } else {
                    $_DGroupName = $_ADGroup.$ADGroupProperty.Name
                }
                $_DGroup = Get-DistributionGroup -Identity $_DGroupName
                if ($_DGroup) {
                    $_pair = @{"pgroup" = $_ADGroup; "dgroup" = $_DGroup}
                    $_pairs.Add($_pair)
                } else {
                    Write-Warning "Unable to find DistributionGroup with name $_DGroup"
                }
            }

        } else {
            #objects
            for ($i = 0; $i -lt $PermissionGroup.Count; $i++) {
                $_pair = @{"pgroup" = $PermissionGroup[$i]; "dgroup" = $DistributionGroup[$i]}
                $_pairs.Add($_pair)
            }
        }

        forEach ($_pair in $_pairs) {
            Write-Verbose $_pair
            
            $_PermissionGroup = $_pair.pgroup
            $_DistributionGroup = $_pair.dgroup

            $Flatten = $true
            if ($DoNotFlatten) { $Flatten = $false}

            #Gather members
            $_PermissionGroupMembers = Get-ADGroupMember -Identity $_PermissionGroup -Recursive:$Flatten | Get-ADUser -Properties "mailNickname"
            $_DistributionGroupMembers = Get-DistributionGroupMember -Identity $_DistributionGroup.Identity | Add-Member -MemberType AliasProperty -Name "mailNickname" -Value "Alias" -PassThru

            #Preform compare
            if (!$ReverseDirection) {
                if (!$_DistributionGroupMembers) {
                    $_Addmember = $_PermissionGroupMembers
                    $_Removemember = @()
                } elseif (!$_PermissionGroupMembers) {
                    $_Removemember = $_DistributionGroupMembers
                    $_Addmember = @()
                } else {
                    $_results = Compare-Object -ReferenceObject $_PermissionGroupMembers -DifferenceObject $_DistributionGroupMembers -Property "mailNickname" -IncludeEqual -PassThru
                    $_Addmember = $_results | where -Property SideIndicator -EQ -Value "<="
                    $_Removemember = $_results | where -Property SideIndicator -EQ -Value "=>"
                }

                #Handle Adds
                if ($_Addmember) {
                    $_Addmember = $_Addmember.mailNickname
                    $_AddMember | Add-DistributionGroupMember -Identity $_DistributionGroup.Identity
                }

                #Handle Removes
                if ($_Removemember) {
                    $_Removemember = $_Removemember.mailNickname
                    $_Removemember | Remove-DistributionGroupMember -Identity $_DistributionGroup.Identity
                }
            } else {
                if (!$_DistributionGroupMembers) {
                    $_Addmember = @()
                    $_Removemember = $_PermissionGroupMembers
                } elseif (!$_PermissionGroupMembers) {
                    $_Removemember = @()
                    $_Addmember = $_DistributionGroupMembers
                } else {
                    $_results = Compare-Object -ReferenceObject $_PermissionGroupMembers -DifferenceObject $_DistributionGroupMembers -Property "mailNickname" -IncludeEqual -PassThru
                    $_Addmember = $_results | where -Property SideIndicator -EQ -Value "=>"
                    $_Removemember = $_results | where -Property SideIndicator -EQ -Value "<="
                }

                #Handle Adds
                if ($_Addmember) {
                    $_Addmember = $_Addmember.mailNickname
                    $_Addmember | Add-ADGroupMember -Identity $_PermissionGroup.DistinguishedName
                }

                #Handle Removes
                if ($_Removemember) {
                    $_Removemember = $_Removemember.mailNickname
                    $_Removemember | Remove-ADGroupMember -Identity $_PermissionGroup.DistinguishedName
                }
            }
        }
	}
}

function Remove-PermissionsDistributionGroup {
	[CmdletBinding(SupportsShouldProcess=$true)]
	Param(
		[Parameter(Mandatory=$true,Position=1,HelpMessage="Mailbox name")] 
            [String]$Name
    )
	Process {
        [String] $_cleanIdentity = $Name.Replace(' ', '_')
        [String] $_PermissionGroup = "EL-$_cleanIdentity"
        [String]$Alias = $_cleanIdentity

        Remove-ADGroup -Identity $_PermissionGroup

        Remove-DistributionGroup -Identity $Alias
    }    
}

Export-ModuleMember -Function "New-SharedMailbox","Sync-SharedMailboxAutoMapping","Add-SharedMailboxGroup", "New-PermissionsDistributionGroup", "Sync-PermissionsDistributionGroup", "Remove-PermissionsDistributionGroup" #TODO "Verb-Noun"