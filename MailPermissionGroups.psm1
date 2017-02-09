
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
    Begin {
        Test-Office365Loaded -ErrorOnFalse
    }
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
	Begin {
        Test-Office365Loaded -ErrorOnFalse
    }
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
	Begin {
        Test-Office365Loaded -ErrorOnFalse
    }
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
	[CmdletBinding(SupportsShouldProcess=$true,DefaultParameterSetName="UseDLName")]
	Param(
		[Parameter(Mandatory=$true,Position=1,HelpMessage="Mailbox name/first part of email address.")] 
            [String]$Name,
        [Parameter(Mandatory=$true,Position=2,HelpMessage="Display name to show in address book.")] 
            [String]$DisplayName,
        [Parameter(Mandatory=$false,Position=3,HelpMessage="Specifies the name fo the AD Group to create.  If not provided, the DLName is used with the optional ADGroupPrefix.")] 
            [String]$ADGroupName,
        [Parameter(Mandatory=$false,HelpMessage="Specifies that the Distribution List Distinguished Name is to be added to a property of the ADGroup that is created.  Uses the property specified by ADGroupProperty or it's default.")]
            [Switch]$UseADGroupProperty,
        [Parameter(Mandatory=$false,HelpMessage="The ADGroup property to use to save the Distribution List Distinguished Name to.  Defaults to info.")] 
            [String]$ADGroupProperty = "info",
        [Parameter(Mandatory=$false, ValueFromPipeline=$true, Position=4,HelpMessage="Optional array of members to add (accepts same objects as Add-ADGroupMember).")] 
            [Object[]] $Members,
        [Parameter(Mandatory=$true,HelpMessage="The OU where the permissions groups will be created.")] 
            [String]$OU = "",
        [Parameter(Mandatory=$false,HelpMessage="A prefix that exists on the ADGroup but not on the DistributionGroup name.  Defaults to 'EL-'.  May be set to empty to have both names match.")] 
            [String]$ADGroupPrefix = "EL-",
        [Parameter(Mandatory=$true,HelpMessage="The primary domain to use for the email address.  Set as primary in the proxy addresses.")] 
            [String]$PrimarySMTPDomain,
        [Parameter(Mandatory=$false,HelpMessage="An array of one or more alternative domain names to setup as proxy addresses.")] 
            [String[]]$OtherSMTPDomain,
        [Parameter(HelpMessage="Returns the new AD Group, defaults to returning the distribution list object.")] 
            [Switch]$ReturnADGroup

	)
    Begin {
        Test-Office365Loaded -ErrorOnFalse
    }	
    Process {
        #Setup names.
        [String] $_cleanIdentity = $Name.Replace(' ', '_')

        if (!$ADGroupName) {
            [String] $_PermissionGroup = "$($ADGroupPrefix)$($_cleanIdentity)"
        } else {
            [String] $_PermissionGroup = $ADGroupName
        }
        [String] $_GroupDescription = "This group is synced with an Exchange Distribution list of similar name."

        [String]$Alias = $_cleanIdentity  #$EmailAddress.Split("@")[0]

        if ($pscmdlet.ShouldProcess("Create AD Group $_PermissionGroupName and Exchange Distribution Group $_DistListName adding members to both.")) {
            #Create the Permission Group
            [Microsoft.ActiveDirectory.Management.ADGroup]$PermissionGroup = New-ADGroup -Path $OU -Name $_PermissionGroup -SamAccountName $_PermissionGroup -Description $_GroupDescription -GroupScope Global -PassThru
            
            Sleep -Seconds 2

            #Add Members to the Permission Group
            if ($Members) {
                Add-ADGroupMember -Identity $PermissionGroup.SamAccountName -Members $Members
            }

            #Create the New-DistributionGroup
            $DistGroup = New-DistributionGroup -Name $DisplayName -Alias $Alias -DisplayName $DisplayName -PrimarySmtpAddress "$($Alias)@$($PrimarySMTPDomain)" -MemberDepartRestriction Closed -MemberJoinRestriction Closed

            #Update PermissionGroup if needed
            if ($UseADGroupProperty) {
                $PermissionGroup | Set-ADGroup -Add @{$($ADGroupProperty)=$DistGroup.DistinguishedName}
            }

            forEach ($other in $OtherSMTPDomain) {
                #TODO - Set-ProxyAddress currently does not support DistributionGroups
                #Set-ProxyAddress -Identity ($DistGroup).Alias -ProxyAddress "$($Alias)@$($other)"
            }

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
        [Parameter(ParameterSetName="UseProperty")]
        [Parameter(ParameterSetName="MatchOnName")]
            [Microsoft.ActiveDirectory.Management.ADGroup[]]$PermissionGroup,
        [Parameter(Mandatory=$true,Position=2,ValueFromPipeline=$true,ParameterSetName="objects",HelpMessage="Exchnage DistributionGroup object")] 
        [Parameter(ParameterSetName="MatchOnName")]
            [PSObject[]]$DistributionGroup,
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ParameterSetName="OU",HelpMessage="The OU where the permissions groups can be found.  All will be synced.")] 
            [String]$OU = "",
        [Parameter(Mandatory=$false,ParameterSetName="OU",HelpMessage="A prefix that exists on the ADGroup but not on the DistributionGroup name.  Defaults to 'EL-'.  Only needed if matching on Name.")] 
        [Parameter(ParameterSetName="MatchOnName")]
            [String]$ADGroupPrefix = "EL-",
        [Parameter(Mandatory=$false,ParameterSetName="OU",HelpMessage="Specifies that the Distribution List Distinguished Name is to be added to a property of the ADGroup that is created.  Uses the property specified by ADGroupProperty or it's default.")]
        [Parameter(ParameterSetName="UseProperty")]
            [Switch]$UseADGroupProperty,
        [Parameter(Mandatory=$false,ParameterSetName="OU",HelpMessage="The ADGroup property to use to find the email distribution group.  Defaults to info.")] 
        [Parameter(ParameterSetName="UseProperty")]
            [String]$ADGroupProperty = "info",
        [Parameter(Mandatory=$false,HelpMessage="Default syncs ADGroup to DistributionGroup, this reverse the sync direction.")] 
            [Switch] $ReverseDirection,
        [Parameter(Mandatory=$false,HelpMessage="Default will get all ADGroup members recursevly and adds them individually, this will add as is.")] 
            [Switch] $DoNotFlatten
	)
	Begin {
        Test-Office365Loaded -ErrorOnFalse
    }
    Process {
        [System.Collections.ArrayList] $_pairs = New-Object System.Collections.ArrayList

        if ($OU) {
            #OU
            if (!$UseADGroupProperty) {   #Use name matching
                $_PermissionGroups = Get-ADGroup -SearchScope OneLevel -SearchBase $OU -Filter "*" | Where-Object -Property Name -Value "$($ADGroupPrefix)*" -like | Get-ADGroup
            } else {   #Use group property
                $_PermissionGroups = Get-ADGroup -SearchScope OneLevel -SearchBase $OU -Filter "*" -Properties $ADGroupProperty | Where-Object -Property $ADGroupProperty -Like -Value "CN*" | Get-ADGroup -Properties $ADGroupProperty
            }

            forEach ($_ADGroup in $_PermissionGroups) {
                if (!$UseADGroupProperty) {
                    if ($ADGroupPrefix) {
                        $_DGroupName = ($_ADGroup.Name).Replace($ADGroupPrefix, "")
                    } else {
                        $_DGroupName = $_ADGroup.Name
                    }
                } else {
                    $_DGroupName = $_ADGroup.$ADGroupProperty
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
            if ($PermissionGroup.Count -gt 0 -and $DistributionGroup.Count -eq 0) {    #Use PermissionGroup only to build pairs
                if ($UseADGroupProperty) { #By using property
                    forEach ($_PGroup in $PermissionGroup) {
                        $_PGroup = $_PGroup | Get-ADGroup -Properties $ADGroupProperty
                        $_DGroup = Get-DistributionGroup -Identity $($_PGroup.$ADGroupProperty)
                        if ($_DGroup) {
                            $_pair = @{"pgroup" = $_PGroup; "dgroup" = $_DGroup}
                            $_pairs.Add($_pair)
                        } else {
                            Write-Verbose "No matching distribution group found for $($_PGroup.Name)"
                        }
                    }
                } else { #Attempt to match by name
                    forEach ($_PGroup in $PermissionGroup) {
                        $_DGroupName = $_PGroup.Name.Replace($ADGroupPrefix, "")
                        $_DGroup = Get-DistributionGroup -Identity $_DGroupName -ErrorAction SilentlyContinue
                        if ($_DGroup) {
                            $_pair = @{"pgroup" = $_PGroup; "dgroup" = $_DGroup}
                            $_pairs.Add($_pair)
                        }
                    }
                }
            } elseif ($DistributionGroup.Count -gt 0 -and $PermissionGroup.Count -eq 0) {  #Use DistributionGroup and name matching to build pairs 
                forEach ($_DGroup in $DistributionGroup) {
                    $_DGroup = $_DGroup | Get-DistributionGroup -ErrorAction SilentlyContinue
                    if ($_DGroup) {
                        $_PGroupName = $ADGroupPrefix + $_DGroup.Alias
                        try {
                            $_PGroup = Get-ADGroup -Identity $_PGroupName -ErrorAction SilentlyContinue
                        } catch {}
                        if ($_PGroup) {
                            $_pair = @{"pgroup" = $_PGroup; "dgroup" = $_DGroup}
                            $_pairs.Add($_pair)
                        }
                    }
                }
            } elseif ($DistributionGroup.Count -eq $PermissionGroup.Count) {   #Use matched objects
                #objects
                for ($i = 0; $i -lt $PermissionGroup.Count; $i++) {
                    $_pair = @{"pgroup" = $PermissionGroup[$i]; "dgroup" = $DistributionGroup[$i]}
                    $_pairs.Add($_pair)
                }
            }else {  #Throw error
                throw("When using both DistributionGroup and PermissionGroup parameters, there must be an equal count of both")
            }
        }

        forEach ($_pair in $_pairs) {
            #Write-Verbose $_pair
            
            $_PermissionGroup = $_pair.pgroup
            $_DistributionGroup = $_pair.dgroup

            $Flatten = $true
            #TODO Fix flatten by fixing the TODO below
            #if ($DoNotFlatten) { $Flatten = $false}

            #Gather members
            #TODO - Only supports USER objects, need to determine type of object returned by Get-ADGroupMember and get the right user.
            $_PermissionGroupMembers = (Get-ADGroupMember -Identity $_PermissionGroup.DistinguishedName -Recursive:$Flatten | Get-ADUser).UserPrincipalName
            $_DistributionGroupMembers = (Get-DistributionGroupMember -Identity $_DistributionGroup.Identity | Where RecipientType -eq UserMailbox | Get-Mailbox).UserPrincipalName

            #Preform compare
            if (!$ReverseDirection) {
                if (!$_DistributionGroupMembers) {
                    $_Addmember = $_PermissionGroupMembers
                    $_Removemember = @()
                } elseif (!$_PermissionGroupMembers) {
                    $_Removemember = $_DistributionGroupMembers
                    $_Addmember = @()
                } else {
                    $_results = Compare-Object -ReferenceObject $_PermissionGroupMembers -DifferenceObject $_DistributionGroupMembers -Property "UserPrincipalName" -IncludeEqual -PassThru
                    $_Addmember = [String[]] ($_results | where -Property SideIndicator -EQ -Value "<=")
                    $_Removemember = [String[]] ($_results | where -Property SideIndicator -EQ -Value "=>")
                }

                #Handle Adds
                if ($_Addmember) {
                    $_Addmember| %{Add-DistributionGroupMember -Identity $_DistributionGroup.Identity -Member (($_ | Get-Mailbox).DistinguishedName)}
                }

                #Handle Removes
                if ($_Removemember) {
                    $_Removemember | %{Remove-DistributionGroupMember -Identity $_DistributionGroup.Identity -Member (($_ | Get-Mailbox).DistinguishedName)}
                }
            } else {
                if (!$_DistributionGroupMembers) {
                    $_Addmember = @()
                    $_Removemember = $_PermissionGroupMembers
                } elseif (!$_PermissionGroupMembers) {
                    $_Removemember = @()
                    $_Addmember = $_DistributionGroupMembers
                } else {
                    $_results = Compare-Object -ReferenceObject $_PermissionGroupMembers -DifferenceObject $_DistributionGroupMembers -Property "UserPrincipalName" -IncludeEqual -PassThru
                    $_Addmember = [String[]] ($_results | where -Property SideIndicator -EQ -Value "=>")
                    $_Removemember = [String[]] ($_results | where -Property SideIndicator -EQ -Value "<=")
                }

                #Handle Adds
                if ($_Addmember) {
                    $_Addmember | %{Get-ADUser -Filter "UserPrincipalName -eq '$($_)'"} | Add-ADGroupMember -Identity $_PermissionGroup.DistinguishedName
                }

                #Handle Removes
                if ($_Removemember) {
                    $_Removemember | %{Get-ADUser -Filter "UserPrincipalName -eq '$($_)'"} | Remove-ADGroupMember -Identity $_PermissionGroup.DistinguishedName
                }
            }
        }
	}
}

function Remove-PermissionsDistributionGroup {
	[CmdletBinding(SupportsShouldProcess=$true,DefaultParameterSetName="objects")]
	Param(
        [Parameter(Mandatory=$true,Position=1,ValueFromPipeline=$true,ParameterSetName="objects",HelpMessage="ADGroup PermissionGroup object")]
        [Parameter(ParameterSetName="UseProperty")]
        [Parameter(ParameterSetName="MatchOnName")]
            [Microsoft.ActiveDirectory.Management.ADGroup[]]$PermissionGroup,
        [Parameter(Mandatory=$true,Position=2,ValueFromPipeline=$true,ParameterSetName="objects",HelpMessage="Exchnage DistributionGroup object")] 
        [Parameter(ParameterSetName="MatchOnName")]
            [PSObject[]]$DistributionGroup,
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ParameterSetName="OU",HelpMessage="The OU where the permissions groups can be found.  All will be synced.")] 
            [String]$OU = "",
        [Parameter(Mandatory=$false,ParameterSetName="OU",HelpMessage="A prefix that exists on the ADGroup but not on the DistributionGroup name.  Defaults to 'EL-'.  Only needed if matching on Name.")] 
        [Parameter(ParameterSetName="MatchOnName")]
            [String]$ADGroupPrefix = "EL-",
        [Parameter(Mandatory=$false,ParameterSetName="OU",HelpMessage="Specifies that the Distribution List Distinguished Name is to be added to a property of the ADGroup that is created.  Uses the property specified by ADGroupProperty or it's default.")]
        [Parameter(ParameterSetName="UseProperty")]
            [Switch]$UseADGroupProperty,
        [Parameter(Mandatory=$false,ParameterSetName="OU",HelpMessage="The ADGroup property to use to find the email distribution group.  Defaults to info.")] 
        [Parameter(ParameterSetName="UseProperty")]
            [String]$ADGroupProperty = "info"
	)
    Begin {
        Test-Office365Loaded -ErrorOnFalse
    }
	Process {
        [System.Collections.ArrayList] $_pairs = New-Object System.Collections.ArrayList

        if ($OU) {
            #OU
            if (!$UseADGroupProperty) {   #Use name matching
                $_PermissionGroups = Get-ADGroup -SearchScope OneLevel -SearchBase $OU -Filter "*" | Where-Object -Property Name -Value "$($ADGroupPrefix)*" -like | Get-ADGroup
            } else {   #Use group property
                $_PermissionGroups = Get-ADGroup -SearchScope OneLevel -SearchBase $OU -Filter "*" -Properties $ADGroupProperty | Where-Object -Property $ADGroupProperty -Like -Value "CN*" | Get-ADGroup -Properties $ADGroupProperty
            }

            forEach ($_ADGroup in $_PermissionGroups) {
                if (!$UseADGroupProperty) {
                    if ($ADGroupPrefix) {
                        $_DGroupName = ($_ADGroup.Name).Replace($ADGroupPrefix, "")
                    } else {
                        $_DGroupName = $_ADGroup.Name
                    }
                } else {
                    $_DGroupName = $_ADGroup.$ADGroupProperty
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
            if ($PermissionGroup.Count -gt 0 -and $DistributionGroup.Count -eq 0) {    #Use PermissionGroup only to build pairs
                if ($UseADGroupProperty) { #By using property
                    forEach ($_PGroup in $PermissionGroup) {
                        $_PGroup = $_PGroup | Get-ADGroup -Properties $ADGroupProperty
                        $_DGroup = Get-DistributionGroup -Identity $($_PGroup.$ADGroupProperty)
                        $_pair = @{"pgroup" = $_PGroup; "dgroup" = $_DGroup}
                        $_pairs.Add($_pair)
                    }
                } else { #Attempt to match by name
                    forEach ($_PGroup in $PermissionGroup) {
                        $_DGroupName = $_PGroup.Name.Replace($ADGroupPrefix, "")
                        $_DGroup = Get-DistributionGroup -Identity $_DGroupName -ErrorAction SilentlyContinue
                        if ($_DGroup) {
                            $_pair = @{"pgroup" = $_PGroup; "dgroup" = $_DGroup}
                            $_pairs.Add($_pair)
                        }
                    }
                }
            } elseif ($DistributionGroup.Count -gt 0 -and $PermissionGroup.Count -eq 0) {  #Use DistributionGroup and name matching to build pairs 
                forEach ($_DGroup in $DistributionGroup) {
                    $_DGroup = $_DGroup | Get-DistributionGroup -ErrorAction SilentlyContinue
                    if ($_DGroup) {
                        $_PGroupName = $ADGroupPrefix + $_DGroup.Alias
                        try {
                            $_PGroup = Get-ADGroup -Identity $_PGroupName -ErrorAction SilentlyContinue
                        } catch {}
                        if ($_PGroup) {
                            $_pair = @{"pgroup" = $_PGroup; "dgroup" = $_DGroup}
                            $_pairs.Add($_pair)
                        }
                    }
                }
            } elseif ($DistributionGroup.Count -eq $PermissionGroup.Count) {   #Use matched objects
                #objects
                for ($i = 0; $i -lt $PermissionGroup.Count; $i++) {
                    $_pair = @{"pgroup" = $PermissionGroup[$i]; "dgroup" = $DistributionGroup[$i]}
                    $_pairs.Add($_pair)
                }
            }else {  #Throw error
                throw("When using both DistributionGroup and PermissionGroup parameters, there must be an equal count of both")
            }
        }

        forEach ($_pair in $_pairs) {
            Write-Verbose "About to remove this ADGroup and Exchange DistributionGroup: $_pair"
            if ($Force -or $pscmdlet.ShouldProcess("Remove the following ADGroup and Exchange DistributionGroup: $_pair ?")) {
                try {
                    $_DGroup = $_pair.dgroup.DistinguishedName | Get-DistributionGroup
                    Remove-DistributionGroup -Identity $_pair.dgroup.DistinguishedName -Confirm:$false

                    $_PGroup = $_pair.pgroup.DistinguishedName | Get-ADGroup
                    Remove-ADGroup -Identity $_pair.pgroup.DistinguishedName -Confirm:$false
                } catch {}                
            }

        }
    }    
}

function Test-Office365Loaded {
    [CmdletBinding(SupportsShouldProcess=$true,DefaultParameterSetName="objects")]
	Param(
        [Switch] $WarningOnFalse,
        [Switch] $ErrorOnFalse
    )
    $warning = ""
    $answer = ((Get-Command Get-DistributionGroup -ErrorAction SilentlyContinue).Count -gt 0)
    if ($answer -eq $false) {
        if ($ErrorOnFalse) {
            Write-Error $warning
            Break
        } elseif ($WarningOnFalse) {
            Write-Warning $warning
        } else {
            return $answer
        }
    } else {
        return $answer
    }
}

Export-ModuleMember -Function "New-SharedMailbox","Sync-SharedMailboxAutoMapping","Add-SharedMailboxGroup", "New-PermissionsDistributionGroup", "Sync-PermissionsDistributionGroup", "Remove-PermissionsDistributionGroup", "Test-Office365Loaded" #TODO "Verb-Noun"