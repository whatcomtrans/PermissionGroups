<#
.SYNOPSIS
TODO

.EXAMPLE
TODO#>


function Add-FolderPermissions {
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
            [String]$Grant="Allow",
        [Parameter(Mandatory=$false,Position=6,HelpMessage="An existing ACL object to modify")]
            $ACLObject
	)
	Begin {
		switch ($Permission) {
	            "RW" {  #Modify shorthand
	                $_FileSystemRights = [System.Security.AccessControl.FileSystemRights] "Modify"
	                break
	            }
	            "RO" {  #ReadAndExecute shorthand
	                $_FileSystemRights = [System.Security.AccessControl.FileSystemRights] "ReadAndExecute"
	                break
	            }
	            "LO" {  #This is used for folders only and makes only the folder appear to the user
	                $_FileSystemRights = [System.Security.AccessControl.FileSystemRights] "Read"
					$InheritanceFlags = "None"
					$PropagationFlags = "None"
	                break
	            }
	            default {
	                $_FileSystemRights = $Permission
	                break
	            }
	        }
			Write-Verbose "Checking for shorthand permissions, setting rights to $_FileSystemRights"
	}
	Process {
        Write-Verbose "Retrieving SID for AssignedTo $AssignedTo"
        try {
            [Microsoft.ActiveDirectory.Management.ADObject] $_object = Get-ADGroup -Identity $AssignedTo -ErrorAction SilentlyContinue
        } catch {
            [Microsoft.ActiveDirectory.Management.ADObject] $_object = Get-ADUser -Identity $AssignedTo -ErrorAction Stop
        }

        Write-Verbose "Testing path $Path"
        if ((Test-Path $Path) -eq $false) {
            throw "Folder $Path not found"
        }

        Write-Verbose "Changing ACL"
        if ($_object) {
		    Write-Verbose "Retrieving ACL for $Path"
		    #[System.Security.AccessControl] $acl = Get-Acl -Path $Path
            if (!$ACLObject) {
                $acl = Get-Acl -Path $Path
            } else {
                $acl = $ACLObject
            }
		    #$acl.SetAccessRuleProtection($True, $False)
		    Write-Verbose "Adding new rule to ACL"
		    [System.Security.AccessControl.FileSystemAccessRule] $rule = New-Object System.Security.AccessControl.FileSystemAccessRule($_object.SID ,$_FileSystemRights, $InheritanceFlags, $PropagationFlags, $Grant)
		    $acl.AddAccessRule($rule)
		    Write-Verbose "Setting ACL for $Path to modified ACL"
		    Set-Acl $Path $acl
            
            return $acl
        } else {
            return $null
        }
	}
}

function Get-FolderPermissionsGroupName {
	[CmdletBinding(SupportsShouldProcess=$true)]
	Param(
		[Parameter(Mandatory=$true,Position=0,HelpMessage="Permission (RW,RO,LO, or a FileSystemRights Enumeration value, see http://msdn.microsoft.com/en-us/library/system.security.accesscontrol.filesystemrights.aspx)")] 
            [String]$Permission,
		[Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=1,HelpMessage="Path to set permission on.")] 
            [String]$Path,
		[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=0,HelpMessage="A hash table of path prefixs to be replaced with a friendly names in the group name.  Often used to replace a share UNC path with a friendly name.  Example, \\server\share = Users")] 
            [System.Collections.Hashtable]$PathCommonNames
	)
	Begin {
        switch ($Permission) {
	            "RW" {  #Modify shorthand
	                $_PermissionAbreviation = "RW"
	                break
	            }
	            "RO" {  #ReadAndExecute shorthand
	                $_PermissionAbreviation = "RO"
	                break
	            }
	            "LO" {  #This is used for folders only and makes only the folder appear to the user
	                $_PermissionAbreviation = "LO"
	                break
	            }
	            default {
	                $_PermissionAbreviation = "SP"  # for special
	                break
	            }
	    }
		Write-Verbose "Checking for shorthand permissions, setting rights to $_PermissionAbreviation"
	}
	Process {
        $_permissiongroup = $Path
        #Calculate group name
        foreach ($_commonPath in $PathCommonNames.keys) {
	        if ($Path.StartsWith($_commonPath) -eq $true) {
		        $_permissiongroup = $PathCommonNames.$_commonPath + "-" + $_permissiongroup.Replace($_commonPath, "")
	        }
        }
        $_permissiongroup = $_permissiongroup.Replace("\\", "\").Replace("\", "-").Replace(" ", "_").Replace(":", "-")
        $_permissiongroup = "FLDR-" + $_permissiongroup + "-" + $_PermissionAbreviation
        return $_permissiongroup
        #need to handle paths greater then X length
	}
}


<#

OUTPUTS
None or Microsoft.ActiveDirectory.Management.ADGroup
        
        Returns the new group object
#>

function New-FolderPermissionsGroup {
	[CmdletBinding(SupportsShouldProcess=$true)]
	Param(
		[Parameter(Mandatory=$true,Position=0,HelpMessage="Permission (RW,RO,LO, or a FileSystemRights Enumeration value, see http://msdn.microsoft.com/en-us/library/system.security.accesscontrol.filesystemrights.aspx)")] 
            [String]$Permission,
		[Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=1,HelpMessage="Path to set permission on.")] 
            [String]$Path,
   		[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=0,HelpMessage="OU (complete path) to place newly created Folder Permissions Groups in.")] 
            [String]$OU,
		[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=0,HelpMessage="A hash table of path prefixs to be replaced with a friendly names in the group name.  Often used to replace a share UNC path with a friendly name.  Example, \\server\share = Users")] 
            [System.Collections.Hashtable]$PathCommonNames
	)
	Process {
        #Calculate group name
        $_groupName = Get-FolderPermissionsGroupName -Permission $Permission -Path $Path -PathCommonNames $PathCommonNames
        #See if group already exists
        try {
            $_group = Get-ADGroup -Identity $_groupName
        } catch {
            #Create group
            [Microsoft.ActiveDirectory.Management.ADGroup] $_group = New-ADGroup -DisplayName $_groupName -SAMAccountName $_groupName -Path $OU -Name $_groupName -GroupCategory Security -Description $Path -GroupScope Global -PassThru
        }

        return $_group
	}
}

function Add-FolderPermissionsGroup {
	[CmdletBinding(SupportsShouldProcess=$true)]
	Param(
		[Parameter(Mandatory=$true,Position=0,HelpMessage="Permission (RW,RO,LO, or a FileSystemRights Enumeration value, see http://msdn.microsoft.com/en-us/library/system.security.accesscontrol.filesystemrights.aspx)")] 
            [String]$Permission,
		[Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=1,HelpMessage="Path to set permission on.")] 
            [String]$Path,
   		[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=0,HelpMessage="OU (complete path) to place newly created Folder Permissions Groups in.")] 
            [String]$OU,
		[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=0,HelpMessage="A hash table of path prefixs to be replaced with a friendly names in the group name.  Often used to replace a share UNC path with a friendly name.  Example, \\server\share = Users")] 
            [System.Collections.Hashtable]$PathCommonNames
	)
	Process {
        #TODO - Test to see if it already exists
        #Create group
        [Microsoft.ActiveDirectory.Management.ADGroup] $_group = New-FolderPermissionsGroup -Permission $Permission -Path $Path -OU $OU -PathCommonNames $PathCommonNames
        #Set folder permissions using group
        Add-FolderPermissions -Permission $Permission -Path $Path -AssignedTo $_group.SAMAccountName
        return $_group
	}
}

function Get-FolderPermissionsGroupOrphans {
	[CmdletBinding(SupportsShouldProcess=$true)]
	Param()
    Process {
        #Get array of Folder Permission Groups
        $_groups = Get-ADGroup -Filter "Name -like 'FLDR-*'" -Properties Description
        foreach ($_group in $_groups) {
            #There are two tests to determine if the group should be removed, if either pass then the group can be deleted
            $_delete = $false
            #Test one - Does the path still exist
            if ($_group.Description) {
                if (!(Test-Path $_group.Description)) {
                    $_delete = $true
                }
            }

            #Test two - Are there any members
            if (($_group | Get-ADGroupMember | Measure-Object).Count -eq 0) {
                $_delete = $true
            }

            if ($_delete) {
                $_group
            }
        }
	}
}

function Update-FolderPermissionGroups {
	[CmdletBinding(SupportsShouldProcess=$true)]
	Param(
		[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=1,HelpMessage="Path to set permission on.")] 
            [String]$Path,
   		[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=2,HelpMessage="OU (complete path) to place newly created Folder Permissions Groups in.")] 
            [String]$OU,
		[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=3,HelpMessage="A hash table of path prefixs to be replaced with a friendly names in the group name.  Often used to replace a share UNC path with a friendly name.  Example, \\server\share = Users")] 
            [System.Collections.Hashtable]$PathCommonNames,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,Position=4,HelpMessage="Recursively process sub-folders?")]
            [Switch] $Recurse,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,Position=5,HelpMessage="Leave the existing rights acl rule in place, otherwise it removes them.")]
            [Switch] $NoRemove
	)
	Process {

        [String[]] $results = @()

        [Boolean] $MadeChange = $false

        # check folder permissions
        $acl = Get-Acl -Path $Path

        # If permissions exist
        $AccessToMove = $acl.Access | where -Property IsInherited -eq -Value $false | where -Property IdentityReference -NE -Value "CREATOR OWNER" | where -Property IdentityReference -NE -Value "BUILTIN\Administrators" | where -Property IdentityReference -NE -Value "NT AUTHORITY\SYSTEM" | where -Property IdentityReference -NotLike "WHATCOMTRANS\FLDR*"
        $AccessFLDR = $acl.Access | where -Property IsInherited -eq -Value $false | where -Property IdentityReference -Like "WHATCOMTRANS\FLDR*"
        if ($AccessFLDR.Count -gt 0) {
            [String []] $ExistingADGroups = ($AccessFLDR.IdentityReference.ToString().Split("\")[1])
        } else {
            [String []] $ExistingADGroups = @()
        }

        #   Get permission type
        ForEach ($Access in $AccessToMove) {
            $Rights = ($Access.FileSystemRights.ToString()).split(",").Trim() | where -FilterScript {$_ -ne "Synchronize"}
            $NeededADGroupName = FolderPermissionsGroupName -Permission (Get-RightsShortName -Right $Rights) -Path $Path -PathCommonNames $PathCommonNames
            if ($access.IdentityReference.Value -eq "BUILTIN\Users") {
                $member = Get-ADGroup "All-Users"    #This is a WTA hack
            } else {
                $LDAPFilter = "(samaccountname=$($Access.IdentityReference.Value.Split("\")[1]))"
                $member = (Get-ADObject -LDAPFilter $LDAPFilter ).DistinguishedName
            }
            if ($NeededADGroupName -icontains $ExistingADGroups) {
                # Move existing permissions into the group
                Add-ADGroupMember -Identity $NeededADGroupName -Members $member
            } else {
                #   If group with permission type does not aleady exist
                #   create a group the permisson

                $Permission = (Get-RightsShortName $Rights)
                [Microsoft.ActiveDirectory.Management.ADGroup] $_group = New-FolderPermissionsGroup -Permission $Permission -Path $Path -OU $OU -PathCommonNames $PathCommonNames
                
                #   Set folder permissions using group
                $acl = Add-FolderPermissions -Permission $Permission -Path $Path -AssignedTo $_group.SAMAccountName -ACLObject $acl

                #   Add member
                $_group | Add-ADGroupMember -Members $member
            }
            
            if (!$NoRemove) {
                #   Remove ACL entry
                #$acl.RemoveAccessRule($Access)           
            }

            $MadeChange = $true
        }
        #Update the ACL if needed
        if ($MadeChange) {
            Set-Acl $Path $acl
            $results += $Path
        }

        if ($Recurse) {
            # For each child folder, call this method recursively
            Get-ChildItem -Path $Path -Directory | %{
                $results += Update-FolderPermissionGroups -Path ($_.FullName) -OU $OU -PathCommonNames $PathCommonNames -Recurse
            }
        }

        return $results
    }
}

function Get-RightsShortName {
    [CmdletBinding(SupportsShouldProcess=$false)]
    param (
        [System.Security.AccessControl.FileSystemRights] $Right
    )
    switch ($Right) {
        "Modify" {
            return "RW"
            break
        }
        "ReadAndExecute" {
            return "RO"
        }
        default {
            return $Right
        }
    }
}


Export-ModuleMember -Function "Add-FolderPermissions", "Get-FolderPermissionsGroupName", "New-FolderPermissionsGroup", "Set-FolderPermissionsOU", "Get-FolderPermissionsOU", "Get-FolderPermissionsGroupOrphans", "Add-FolderPermissionsGroup", "Update-FolderPermissionGroups"