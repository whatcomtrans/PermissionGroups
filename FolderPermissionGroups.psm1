﻿<#
.SYNOPSIS
TODO

.EXAMPLE
TODO#>

#Shared variables
$_shareCommonNames = @{}
$_FolderPermissionsOU = ""

#Remove once everything is working with parameter value defaults
#$_shareCommonNames = @{"\\wtafx\public" = "Public"; "\\wtafx\restricted" = "Restricted"; "\\wtafx\users" = "Users"; "\\wtafx\applications" = "Apps"}
#$_FolderPermissionsOU = "OU=FolderPermissionGroups,OU=PermissionGroups,DC=whatcomtrans,DC=net"

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
            [String]$Grant="Allow"
	)
	Begin {
        #Load any ParameterDefaultValues
        Set-FolderPermissionGroupSettings

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
            $acl = Get-Acl -Path $Path
		    #$acl.SetAccessRuleProtection($True, $False)
		    Write-Verbose "Adding new rule to ACL"
		    [System.Security.AccessControl.FileSystemAccessRule] $rule = New-Object System.Security.AccessControl.FileSystemAccessRule($_object.SID ,$_FileSystemRights, $InheritanceFlags, $PropagationFlags, $Grant)
		    $acl.AddAccessRule($rule)
		    Write-Verbose "Setting ACL for $Path to modified ACL"
		    Set-Acl $Path $acl
        } else {
            
        }
	}
}

function Get-FolderPermissionsGroupName {
	[CmdletBinding(SupportsShouldProcess=$true)]
	Param(
		[Parameter(Mandatory=$true,Position=0,HelpMessage="Permission (RW,RO,LO, or a FileSystemRights Enumeration value, see http://msdn.microsoft.com/en-us/library/system.security.accesscontrol.filesystemrights.aspx)")] 
            [String]$Permission,
		[Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=1,HelpMessage="Path to set permission on.")] 
            [String]$Path
	)
	Begin {
        #Load any ParameterDefaultValues
        Set-FolderPermissionGroupSettings

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
        foreach ($_commonPath in $_shareCommonNames.keys) {
	        if ($Path.StartsWith($_commonPath) -eq $true) {
		        $_permissiongroup = $_shareCommonNames.$_commonPath + $_permissiongroup.Replace($_commonPath, "")
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
            [String]$Path
	)
    Begin {
        #Load any ParameterDefaultValues
        Set-FolderPermissionGroupSettings
    }
	Process {
        #Calculate group name
        $_groupName = Get-FolderPermissionsGroupName -Permission $Permission -Path $Path
        #See if group already exists
        try {
            $_group = Get-ADGroup -Identity $_groupName
        } catch {
            #Create group
            [Microsoft.ActiveDirectory.Management.ADGroup] $_group = New-ADGroup -DisplayName $_groupName -SAMAccountName $_groupName -Path $_FolderPermissionsOU -Name $_groupName -GroupCategory Security -Description $Path -GroupScope Global -PassThru
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
            [String]$Path
	)
    Begin {
        #Load any ParameterDefaultValues
        Set-FolderPermissionGroupSettings
    }
	Process {
        #TODO - Test to see if it already exists
        #Create group
        [Microsoft.ActiveDirectory.Management.ADGroup] $_group = New-FolderPermissionsGroup -Permission $Permission -Path $Path
        #Set folder permissions using group
        Add-FolderPermissions -Permission $Permission -Path $Path -AssignedTo $_group.SAMAccountName
        return $_group
	}
}

function Set-FolderPermissionGroupSettings {
    [CmdletBinding(SupportsShouldProcess=$true)]
	Param(
		[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=0,HelpMessage="OU (complete path) to place newly created Folder Permissions Groups in.")] 
            [String]$OU,
		[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=0,HelpMessage="A hash table of path prefixs to be replaced with a friendly names in the group name.  Often used to replace a share UNC path with a friendly name.  Example, \\server\share = Users")] 
            [System.Collections.Hashtable]$PathCommonNames
    )
    if ($OU) {
        Set-FolderPermissionsOU $OU
    }
    if ($PathCommonNames) {
        Set-PathCommonNames $PathCommonNames
    }
}

function Get-FolderPermissionsOU {
    [CmdletBinding(SupportsShouldProcess=$false)]
	Param()
    return $_FolderPermissionsOU
}

function Set-FolderPermissionsOU {
    [CmdletBinding(SupportsShouldProcess=$true)]
	Param(
		[Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0,HelpMessage="OU (complete path) to place newly created Folder Permissions Groups in.")] 
            [String]$OU
	)
    $_FolderPermissionsOU = $OU
}

function Get-PathCommonNames {
    [CmdletBinding(SupportsShouldProcess=$false)]
	Param()
    return $_shareCommonNames
}

function Set-PathCommonNames {
    [CmdletBinding(SupportsShouldProcess=$true,DefaultParameterSetName="WholeHash")]
	Param(
		[Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0,ParameterSetName="WholeHash",HelpMessage="A hash table of path prefixs to be replaced with a friendly names in the group name.  Often used to replace a share UNC path with a friendly name.  Example, \\server\share = Users")] 
            [System.Collections.Hashtable]$PathCommonNames,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,Position=1,ParameterSetName="WholeHash",HelpMessage="Replace whole hash table with new one.")] 
            [Switch]$Replace,
        [Parameter(Mandatory=$true,ValueFromPipeline=$false,Position=0,ParameterSetName="Add",HelpMessage="Used to easily add an individual entry to the hash tabel")]
            [switch]$Add,
        [Parameter(Mandatory=$true,ValueFromPipeline=$false,Position=1,ParameterSetName="Add",HelpMessage="The path prefix to add.")] 
            [String]$AddPath,
        [Parameter(Mandatory=$true,ValueFromPipeline=$false,Position=2,ParameterSetName="Add",HelpMessage="The friendly name to replace the path prefix with in the group name.")] 
            [String]$CommonName,
        [Parameter(Mandatory=$true,ValueFromPipeline=$false,Position=0,ParameterSetName="Remove",HelpMessage="Used to easily add an individual entry to the hash tabel")]
            [switch]$Remove,
        [Parameter(Mandatory=$true,ValueFromPipeline=$false,Position=1,ParameterSetName="Remove",HelpMessage="The path prefix to add.")] 
            [String]$RemovePath
	)
    if ($Replace) {
        $_shareCommonNames = $ShareCommonNames
    }
    if ($Add) {
        $_shareCommonNames.Add($AddPath, $CommonName)
    }
    if ($Remove) {
        $_shareCommonNames.Remove($RemovePath)
    }
}


function Get-FolderPermissionsGroupOrphans {
	[CmdletBinding(SupportsShouldProcess=$true)]
	Param()
	Begin {
        #Load any ParameterDefaultValues
        Set-FolderPermissionGroupSettings
    }
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

Export-ModuleMember -Function "Add-FolderPermissions", "Get-FolderPermissionsGroupName", "New-FolderPermissionsGroup", "Set-FolderPermissionsOU", "Get-FolderPermissionsOU", "Get-FolderPermissionsGroupOrphans", "Add-FolderPermissionsGroup", "Get-PathCommonNames", "Set-PathCommonNames", "Set-FolderPermissionGroupSettings"