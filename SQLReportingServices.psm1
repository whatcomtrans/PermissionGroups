<#
.SYNOPSIS
A template for cmdlets

#>
function New-SRSPermissionsGroup {
	[CmdletBinding(SupportsShouldProcess=$true)]
	Param(
		[Parameter(Mandatory=$true,ValueFromPipeline=$true,HelpMessage="put help here")]
		[String]$Folder,
        [Parameter(HelpMessage="put help here")]
        [String]$Permission="Browser",
        [Parameter(HelpMessage="put help here")]
        [String]$PermissionsGroupOU="OU=SQLSRSPermissionGroups,OU=PermissionGroups,DC=whatcomtrans,DC=net"
	)
	Process {
        $FolderName = $Folder.Replace("\", "-").Replace(" ", "_")
        $Permission = $Permission.Replace(" ", "_")

        $GroupName = "SRS-$FolderName-$Permission"
        $Description = "A PermissionsGroups for SQL Report Writing Services with $Permission to $Folder"
        $_newGroup = New-ADGroup -Name $GroupName -DisplayName $GroupName -Description $Description -GroupCategory Security -GroupScope Global -PassThru -SamAccountName $GroupName -Path $PermissionsGroupOU
        return $_newGroup
	}
}

Export-ModuleMember -Cmdlet New-SRSPermissionsGroup