<#
.SYNOPSIS
Opens a GridView to select employee AD accounts from.

#>
function Choose-EmployeeADUser {
	[CmdletBinding(SupportsShouldProcess=$false)]
	Param(
		[Parameter(Mandatory=$false,Position=0,ValueFromPipeline=$false,HelpMessage="Defaults to OutputMode Multiple, use this switch to change to Single")]
		[Switch] $Single,
        [Parameter(Mandatory=$false,Position=1,ValueFromPipeline=$false,HelpMessage="OU to search for employess in")]
        [String] $SearchBase
	)
	Process {
        if ($Single) {
            (Get-ADUser -Filter {DistinguishedName -like "*"} -SearchBase $SearchBase -Properties Title,Department | Out-GridView -OutputMode Single -Passthrough -Title "Choose Employee(s)")
        } else {
            (Get-ADUser -Filter {DistinguishedName -like "*"} -SearchBase $SearchBase -Properties Title,Department | Out-GridView -OutputMode Multiple -Title "Choose Employee(s)")
        }
	}
}

function Get-RoleADGroup {
	[CmdletBinding(SupportsShouldProcess=$false,DefaultParameterSetName="example")]
	Param(
        [Parameter(Mandatory=$true,Position=0,HelpMessage="The name of the Role Type, must match the name of the OU containing the groups requested")]
		[String]$RoleType,
		[Parameter(Mandatory=$false,HelpMessage="ONLY USE IN ISE, pipes results to Out-GridView for further filtering before returning results")]
        [Switch]$ShowGridView,
		[Parameter(Mandatory=$true,HelpMessage="The base part of the OU that when appended to RoleType forms the SearchBase for the Get-ADGroup command")]
		[String]$OUBase
	)
	Process {
        $groups = Get-ADGroup -Filter * -SearchBase "OU=$($RoleType),$($OUBase)"
        if ($ShowGridView) {
            return $groups | Out-GridView -OutputMode Multiple -Title "$RoleType Role Groups"
        } else {
            return $groups
        }
	}
}

<#
This cmdlet is used for changing a user's role.  It ASSUMES that as user should have only one ROLE of the TYPE (OU) provided in the NewRole parameter.
The function identifies the OU of the NewRole AD Group.  It then adds the new group to the Identity user as a Member Of

AND

removes all other AD groups from the user Member Of that match the same OU as the NewRole AD Group.
#>
function Swap-RoleADGroup {
	[CmdletBinding(SupportsShouldProcess=$true)]
	Param(
        [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Identity of User to swap roles on.")]
		[String]$Identity,
        [Parameter(Mandatory=$true,Position=1,ValueFromPipelineByPropertyName=$true,HelpMessage="The Identity of the new role group to assign the user to.")]
		[String]$NewRole
	)
	Process {
        $NewRoleGroup = Get-ADGroup -Identity $NewRole
        $New
	}
}



function Get-RolePositionGroup {
    [CmdletBinding(SupportsShouldProcess=$false,DefaultParameterSetName="OU")]
	Param(
		[Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,HelpMessage="The title of the position which we are returning the matching role group")]
		[Object]$PositionTitle,
        [Parameter(Mandatory=$false,ParameterSetName="OU",HelpMessage="Active Directory OU to limit search of role group to")]
        [String]$OU,
        [Parameter(Mandatory=$false,HelpMessage="The property to match PositionTitle against")]
        [String]$PropertyName="displayName"
	)
	Process {
        if ($OU) {
            return Get-ADGroup -Filter "$PropertyName -EQ '$PositionTitle'" -SearchBase $OU
        } else {
            return Get-ADGroup -Filter "$PropertyName -EQ '$PositionTitle'"
        }
	}
}

function Get-ADGroupMemberChanges {
	[CmdletBinding(SupportsShouldProcess=$false)]
	Param(
		[Parameter(Mandatory=$true,Position=0,ValueFromPipelineByPropertyName=$true,HelpMessage="Group to compare")]
		[String] $SamAccountName,
		[Parameter(Mandatory=$false,Position=1,HelpMessage="Path to store comparison data, defaults to current directory")]
        [String] $Path = ".\",
        [Parameter(Mandatory=$false,HelpMessage="Do not save the changes")]
        [Switch] $NoSave
	)
	Begin {
        #Put begining stuff here
	}
	Process {
        [String] $filePath = "$($Path)$($SamAccountName).xml"
        if (Test-Path $filePath) {
            $newValues = (Get-ADGroupMember -Identity $SamAccountName -Recursive).SID.Value
            $results = Compare-Object -ReferenceObject (Import-Clixml $filePath) -DifferenceObject $newValues | Select-Object -Property @{Name = "SID"; Expression = {$_.InputObject}},@{Name = "Change"; Expression = {if ($_.SideIndicator -eq "=>") {return "Add"} else {return "Remove"}}}
            if (!$NoSave) {
                $newValues | Export-Clixml $filePath
            }
            $asHashTable = @{Add = ($results | Where-Object -Property Change -EQ -Value Add | ForEach-Object {Get-ADUser -Identity $_.SID}); Remove = ($results | Where-Object -Property Change -EQ -Value Remove | ForEach-Object {Get-ADUser -Identity $_.SID})}
            return $asHashTable
        } else {
            #Initial data set
            (Get-ADGroupMember -Identity $SamAccountName -Recursive).SID.Value | Export-Clixml $filePath
            Write-Warning -Message "Initial load of data set, no comparison made"
            return $Null
        }
	}
	End {
        #Put end here
	}
}

function Test-ADGroup {
	[CmdletBinding(SupportsShouldProcess=$false)]
	Param(
		[Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
		[String]$Identity
	)
	Begin {
        #Put begining stuff here
	}
	Process {
        #Put process here
        $_groupExists = $false
        try {
            if (Get-ADGroup $Identity) {
                $_groupExists = $true
            }
        } catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
            $_groupExists = $false
        }
        return $_groupExists
	}
	End {
        #Put end here
	}
}


<#
.SYNOPSIS
TODO

#>
function Sync-ADGroupExpanded {
	[CmdletBinding(SupportsShouldProcess=$true,DefaultParameterSetName="group")]
	Param(
		[Parameter(Mandatory=$false,Position=0,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,ParameterSetName="group")]
		[PSObject]$ExpandedGroup,
        [Parameter(Mandatory=$false,Position=1,ValueFromPipeline=$false,ValueFromPipelineByPropertyName=$true,ParameterSetName="group")]
        [PSObject]$CompactGroup,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,ValueFromPipelineByPropertyName=$true,ParameterSetName="OU")]
        [String]$OU,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,ValueFromPipelineByPropertyName=$true)]
        [String]$Prefix="",
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,ValueFromPipelineByPropertyName=$true)]
        [String]$Postfix = "-Expanded",
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,ValueFromPipelineByPropertyName=$false)]
        [Switch]$ReturnIssueFix
	)
	Begin {
        #Put begining stuff here
	}
	Process {
        #Put process here
        $skip = $false
        If ($OU) {
            #Get all matching groups from an OU and recusively expand them
            $_filter = "Name -like '$($Prefix)*$($Postfix)'"
            $ExpandedGroups = Get-ADGroup -Filter $_filter -SearchBase $OU -SearchScope Subtree
            if ($ExpandedGroups) {
                $ExpandedGroups | Sync-ADGroupExpanded -Prefix $Prefix -Postfix $Postfix -ReturnIssueFix:$ReturnIssueFix
            } else {
                Write-Warning "No matching groups found in OU"
            }
        } else  {
            Try {
                if (!$ExpandedGroup) {
                    if ($CompactGroup) {
                        $_ExpandedGroupIdentity = "$($Prefix)$($CompactGroup.SamAccountName)$($Postfix)"
                        $ExpandedGroup = Get-ADGroup -Identity $_ExpandedGroupIdentity
                    } else {
                        Write-Warning "Must specify either ExpanedGroup or CompactGroup"
                        $skip = $true
                    }
                } elseif (!$CompactGroup) {
                    if ($ExpandedGroup) {
                        $_CompactGroupIdentity = $ExpandedGroup.SamAccountName
                        If ($Prefix) {
                            $_CompactGroupIdentity = $_CompactGroupIdentity.Replace($Prefix, "")
                        }
                        If ($Postfix) {
                            $_CompactGroupIdentity = $_CompactGroupIdentity.Replace($Postfix, "")
                        }
                        $CompactGroup = Get-ADGroup -Identity $_CompactGroupIdentity
                    } else {
                        Write-Warning "Must specify either ExpanedGroup or CompactGroup"
                        $skip = $true
                    }
                }
            } catch {
                Write-Warning "Invalid Groups specified"
                $skip = $true
            } 

            if (!$skip) {
    
                $CompactGroupMembers = Get-ADGroupMember -Identity $CompactGroup.DistinguishedName -Recursive
    
                $ExpandedGroupMembers = Get-ADGroupMember -Identity $ExpandedGroup.DistinguishedName

                $removeMembers = @()
                $addMembers = @()
    
                if (!$ExpandedGroupMembers) {
                    $addMembers = $CompactGroupMembers
                } elseif (!$CompactGroupMembers) {
                    $removeMembers = $ExpandedGroupMembers
                } else {
                    $changes = Compare-Object -ReferenceObject $CompactGroupMembers -DifferenceObject $ExpandedGroupMembers
                    $removeMembers = ($changes | Where-Object SideIndicator -EQ "=>").InputObject
                    $addMembers = ($changes | Where-Object SideIndicator -EQ "<=").InputObject
                }
    
                #Remove members
                if ($removeMembers) {
                    if ($ReturnIssueFix) {
                        ForEach ($removeMember in $removeMembers) {
                            Write-Output (New-IssueFix -FixCommandString "Remove-ADGroupMember -Identity '$($ExpandedGroup.DistinguishedName)' -Members '$($removeMember.DistinguishedName)'" -FixDescription "Remove $($removeMember.SamAccountName) from EXPANDED group: $($ExpandedGroup.SamAccountName)")
                        }
                    } else {
                        Remove-ADGroupMember -Identity $ExpandedGroup.DistinguishedName -Members $removeMembers
                    }
                }

                #Add members
                if ($addMembers) {
                    if ($ReturnIssueFix) {
                        ForEach ($addMember in $addMembers) {
                            Write-Output (New-IssueFix -FixCommandString "Add-ADGroupMember -Identity '$($ExpandedGroup.DistinguishedName)' -Members '$($addMember.DistinguishedName)'" -FixDescription "Add $($addMember.SamAccountName) to EXPANDED group: $($ExpandedGroup.SamAccountName)")
                        }
                    } else {
                        Add-ADGroupMember -Identity $ExpandedGroup.DistinguishedName -Members $addMembers
                    }
                }
            }    
        }
	}
	End {
        #Put end here
	}
}

<#
.SYNOPSIS
Removes from a specific user all groups from MemberOf.
Default leaves Domain Users but use the parameter ExceptGroup to change the group to leave or set to empty string to remove all groups.

#>
function Remove-ADUserMemberships {
	[CmdletBinding(SupportsShouldProcess=$true,DefaultParameterSetName="user")]
	Param(
		[Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$false,ParameterSetName="user")]
		[Object] $Identity,
        [Parameter(Mandatory=$false,Position=1,ValueFromPipeline=$false,ValueFromPipelineByPropertyName=$false,ParameterSetName="user")]
        [String] $ExceptGroup = "Domain Users",
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,ValueFromPipelineByPropertyName=$false)]
        [Switch]$ReturnIssueFix
	)
	Process {
        $alias = ($Identity | Get-ADUser).DistinguishedName
        Get-ADPrincipalGroupMembership -Identity $alias| where {$_.Name -notlike $ExceptGroup} |% {
            if ($ReturnIssueFix) {
                Write-Output (New-IssueFix -FixCommandString "Remove-ADPrincipalGroupMembership -Identity '$alias' -MemberOf '$($_.DistinguishedName)'" -FixDescription "Removing from user $alias group $($_.samaccountname)")
            } else {
                Remove-ADPrincipalGroupMembership -Identity $alias -MemberOf $_
            }
        }
	}
}

function Move-ADGroupsFromUsersToGroup {
	[CmdletBinding(SupportsShouldProcess=$true,DefaultParameterSetName="objects")]
	Param(
		[Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ParameterSetName="objects")]
		[Microsoft.ActiveDirectory.Management.ADGroup] $Group,
                [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,ParameterSetName="names")]
                [String] $Name,
                [Parameter(Mandatory=$false,Position=1)]
                [Microsoft.ActiveDirectory.Management.ADGroup[]] $GroupsToMove
	)
	Begin {
        #Put begining stuff here
	}
	Process {
        #If just passed name, get group
        if (!$Group) {
            $Group = Get-ADGroup -Identity $Name
        }

        [Microsoft.ActiveDirectory.Management.ADPrincipal[]] $Users = Get-ADGroupMember -Identity $Group

        #If not passed a list of groups to move, ask for them from the first user
        if (!$GroupsToMove) {
            $GroupsToMove = Get-ADPrincipalGroupMembership -Identity $Users[0] | Out-GridView -Title "Select Groups from $($Users[0].Name) to Move from Users to $($Group.Name)" -OutputMode Multiple
        }

        $GroupsToMove | %{Add-ADGroupMember -Identity $_ -Members $Group}
        $GroupsToMove | %{Remove-ADGroupMember -Identity $_ -Members $Users}

	}
	End {
        #Put end here
	}
}

function Set-ADUserPrimaryGroup {
	[CmdletBinding(SupportsShouldProcess=$true,DefaultParameterSetName="user")]
	Param(
		[Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$false,ParameterSetName="user")]
		[Object] $Identity,
        [Parameter(Mandatory=$true,Position=1,ValueFromPipeline=$false,ValueFromPipelineByPropertyName=$false,ParameterSetName="user")]
        [Object] $Group,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,ValueFromPipelineByPropertyName=$false)]
        [Switch]$ReturnIssueFix
	)
	Process {
        $Identity = ($Identity | Get-ADUser -Properties *)
        $Group = ($Group | Get-ADGroup)

        #Verify it is not already their primary group
        if ($Identity.PrimaryGroup -ne $Group.DistinguishedName) {

            #Verify group is already a member, if not, add it
            if ($Identity.MemberOf -notcontains $Group.DistinguishedName) {
                Add-ADPrincipalGroupMembership -Identity $Identity.DistinguishedName -MemberOf $Group
            }

            $groupSID = $Group.SID
            $groupRID = $groupSID.Value.Replace($groupSID.AccountDomainSid, "").Replace("-", "")

            if ($ReturnIssueFix) {
                Write-Output (New-IssueFix -FixCommandString "Set-ADObject -Identity '$($Identity.DistinguishedName)' -Replace @{PrimaryGroupID=$groupRID}" -FixDescription "For user $($Identity.SamAccountName) setting primary group to $($Group.samaccountname)")
            } else {
                Set-ADObject -Identity $Identity.DistinguishedName -Replace @{PrimaryGroupID=$groupRID}
            }
        }
	}
}

function Get-GroupMembershipRecursive {
    [CmdletBinding(SupportsShouldProcess=$false,DefaultParameterSetName="example")]
	Param(
		[Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,HelpMessage="Identity of Group, same as Get-ADPrincipalGroupMembership")]
		[Object] $Identity,
        [Parameter(Mandatory=$false,Position=1,ValueFromPipeline=$false,HelpMessage="AD Server to use.  Make sure it IsGlobalCatalog and works with Get-ADPrincipalGroupMembership")]
        [String] $Server
	)
    Process {
        $retGroups = @()
        $groups = Get-ADPrincipalGroupMembership -Identity $Identity -Server $Server
        $retGroups += $groups
        if ($groups) {
            foreach ($group in $groups) {
                $retGroups += (Get-GroupMembershipRecursive -Identity $group -Server $Server)
            }
        }
        return $retGroups
    }
}

function Copy-ADGroup {
    [CmdletBinding(SupportsShouldProcess=$false,DefaultParameterSetName="example")]
	Param(
		[Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,HelpMessage="Identity of Group, same as Get-ADPrincipalGroupMembership")]
        [Object]$Identity,
        [Parameter(Mandatory=$true,Position=1,ValueFromPipeline=$false,HelpMessage="New Identity of Group, same as Get-ADPrincipalGroupMembership")]
        [String]$NewName,
        [Parameter(Mandatory=$false,Position=2,ValueFromPipeline=$false,HelpMessage="Skip copy of group members")]
        [Switch]$SkipMembers,
        [Parameter(Mandatory=$false,Position=3,ValueFromPipeline=$false,HelpMessage="Skip copy of MemberOf")]
        [Switch]$SkipMembersOf
	)
    Process {
        #Get group to copy
        $toCopy = Get-ADGroup -Identity $Identity -Properties *
        $OUpath = (([adsi]"LDAP://$($toCopy.DistinguishedName)").Parent).Substring(7)

        #Create new group, using the toCopy as a template
        $newGroup = Get-ADGroup -Identity $Identity -Properties * | New-ADGroup -Name $NewName -SamAccountName $NewName -Path $OUpath -PassThru

        Start-Sleep -Seconds 5

        $newGroup = Get-ADGroup -Identity $newGroup.DistinguishedName -Properties *

        #Copy all Group Members
        if (!$SkipMembers) {
            Add-ADGroupMember -Identity $newGroup.DistinguishedName -Members $toCopy.Members
        }

        #Copy all Group MemberOf
        if (!$SkipMemberOf) {
            Add-ADPrincipalGroupMembership -Identity $newGroup.DistinguishedName -MemberOf $toCopy.MemberOf
        }

        Write-Output $newGroup
    }
}

Export-ModuleMember -Function *