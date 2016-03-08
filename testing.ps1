Import-Module PShould
$Num = 6

Import-Module .\PermissionGroups -Force -Verbose

#Create new Email Distribution Group linked to an AD "permissions" Security group
New-PermissionsDistributionGroup -Name "TestDL$Num" -DisplayName "Test Email Distribution List $Num" -UseADGroupProperty -OU "OU=Email,OU=PermissionGroups,DC=whatcomtrans,DC=net" -PrimarySMTPDomain "ridewta.com" -OtherSMTPDomain @("whatcomtrans.net") -ReturnADGroup
#Sync-PermissionsDistributionGroup -PermissionGroup (Get-ADGroup "EL-TestDL2") -UseADGroupProperty -ReverseDirection -DoNotFlatten #-OU "OU=Email,OU=PermissionGroups,DC=whatcomtrans,DC=net" | should count 1
#Sync-PermissionsDistributionGroup -PermissionGroup (Get-ADGroup "EL-TestDL2") -ADGroupPrefix "EL-"
#Sync-PermissionsDistributionGroup -DistributionGroup "TestDL2" -ADGroupPrefix "EL-"

#Remove-PermissionsDistributionGroup -PermissionGroup (Get-ADGroup "EL-TestDL2") -UseADGroupProperty

#help Sync-PermissionsDistributionGroup