
#Create the Containing folder Projects
$dirPath = "\\wtafx\restricted"
$dirPath = New-Item -Path $dirPath -Name "Projects" -ItemType directory
$_folderGroupRoot = Add-FolderPermissionsGroup -Permission LO -Path ($dirPath.Fullname)


#Create WMS project

#Create Project ROLE Group
$groupName = "WMSProject"
$_ProjectsRolePermissionsOU = "OU=Projects,OU=RoleGroups,DC=whatcomtrans,DC=net"
$groupDisplayName = "Workforce Management System Project Team"
[Microsoft.ActiveDirectory.Management.ADGroup] $_group = New-ADGroup -DisplayName $groupDisplayName -SAMAccountName $groupName -Path $_ProjectsRolePermissionsOU -Name $groupName -GroupCategory Security -Description "" -GroupScope Global -PassThru
Enable-SecurityGroupAsDistributionGroup -Identity $_group -DisplayName $groupDisplayName -EmailAddress "$_groupName@ridewta.com"

#Create folder
$dirPath = New-Item -Path ($dirPath.FullName) -Name $groupName -ItemType directory
$_folderGroupProject = Add-FolderPermissionsGroup -Permission RW -Path ($dirPath.FullName)
sleep 20
Add-ADGroupMember -Identity $_folderGroupProject -Members $_group
Add-ADGroupMember -Identity $_folderGroupRoot -Members $_folderGroupProject