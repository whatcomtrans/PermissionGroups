$permission_group = Add-FolderPermissionsGroup -Permission RW -Path "\\wtafx\restricted\IT"

Add-ADGroupMember -Identity $permission_group -Members (Get-ADUser -Filter {Name -Like '*'} | OUt-GridView -OutputMode Multiple)
Add-ADGroupMember -Identity $permission_group -Members (Get-ADGroup -Filter {Name -Like '*'} | OUt-GridView -OutputMode Multiple)

<#
$permission_group = Add-FolderPermissionsGroup -Permission RO -Path "\\wtafx\restricted\OpsSup\Accidents &  Incidents"
$permission_group = Get-FolderPermissionsGroupName -Permission RO -Path "\\wtafx\restricted\OpsSup\Accidents &  Incidents"

"FLDR-Restricted-OpsSup-Accidents_&__Incidents-RO"
#>