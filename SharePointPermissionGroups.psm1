function New-SharePointPermissionsGroup {
	[CmdletBinding(SupportsShouldProcess=$true)]
	Param(
		[Parameter(Mandatory=$true,Position=0,HelpMessage="Permission (RW,RO,ED,DE,FC)")] 
            [String]$Permission,
		[Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=1,HelpMessage="Subsite to set permission on.")] 
            [String]$Subsite,
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=2,HelpMessage="Site to set permission on.")] 
            [String]$Site,
   		[Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=3,HelpMessage="OU (complete path) to place newly created Folder Permissions Groups in.")] 
            [String]$OU
	)
	Process {
        switch ($Permission) {
	        "RW" {
	            $SPOPermission = "Contribute"
	            break
	        }
            "ED" {
	            $SPOPermission = "Edit"
	            break
	        }
	        "RO" {
	            $SPOPermission = "Read"
	            break
	        }
            "DE" {
                $SPOPermission = "Design"
	            break
	        }
	        "FC" {
                $SPOPermission = "Full Control"
	            break
	        }
        }

        $_groupName = "SHPR-$Subsite-$Permission"

        #Create AD permission group
        [Microsoft.ActiveDirectory.Management.ADGroup] $_group = New-ADGroup -DisplayName $_groupName -SAMAccountName $_groupName -Path $OU -Name $_groupName -GroupCategory Security -Description "SharePoint $Site site permission $SPOPermission" -GroupScope Global -PassThru
        Sleep 10
        Force-DirSync
        Sleep 10
    }
}
