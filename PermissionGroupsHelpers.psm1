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
        [String]$Postfix = "-Expanded"
	)
	Begin {
        #Put begining stuff here
	}
	Process {
        #Put process here
        $skip = $false
        If ($OU) {
            #TODO - Get all matching groups from an OU and recusively expand them
            $_filter = "Name -like '$($Prefix)*$($Postfix)'"
            $ExpandedGroups = Get-ADGroup -Filter $_filter -SearchBase $OU -SearchScope Subtree
            if ($ExpandedGroups) {
                $ExpandedGroups | Sync-ADGroupExpanded -Prefix $Prefix -Postfix $Postfix
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
                    Remove-ADGroupMember -Identity $ExpandedGroup.DistinguishedName -Members $removeMembers
                }
                #Add members
                if ($addMembers) {
                    Add-ADGroupMember -Identity $ExpandedGroup.DistinguishedName -Members $addMembers
                }
            }    
        }
	}
	End {
        #Put end here
	}
}

Export-ModuleMember -Function *