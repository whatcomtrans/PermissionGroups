Import-Module PowerShellGet

#Publish to PSGallery and install/import locally

Publish-Module -Path .\ -Repository PSGallery -Verbose
Install-Module -Name PermissionGroups -Repository PSGallery -Force
Import-Module -Name PermissonGroups -Force