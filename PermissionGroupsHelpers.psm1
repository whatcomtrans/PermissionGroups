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
            (Get-ADUser -Filter {DistinguishedName -like "*"} -SearchBase "OU=Employees,DC=whatcomtrans,DC=net" | Out-GridView -OutputMode Single -Passthrough -Title "Choose Employee(s)")
        } else {
            (Get-ADUser -Filter {DistinguishedName -like "*"} -SearchBase "OU=Employees,DC=whatcomtrans,DC=net" | Out-GridView -OutputMode Multiple -Title "Choose Employee(s)")
        }
	}
}

Export-ModuleMember -Function "Choose-EmployeeADUser"