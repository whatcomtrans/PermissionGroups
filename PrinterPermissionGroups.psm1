function New-PrinterGroup {
	[CmdletBinding(SupportsShouldProcess=$true)]
	Param(
		[Parameter(Mandatory=$true,Position=1,ValueFromPipelineByPropertyName=$True,HelpMessage="The name of the printer as returned by Get-Printer.  May pass in result(s) of Get-Printer")] 
            [String]$Name,
        [Parameter(Mandatory=$true,Position=2,ValueFromPipelineByPropertyName=$True,HelpMessage="The name of the print server.  Can come from a printer object passed to cmdlet.")] 
            [String]$ComputerName,
		[Parameter(Mandatory=$false,Position=3,HelpMessage="Turn on(true)/off(false) setting the printer permissions to this new group and only this group, defaults to True")] 
            [Switch]$SetSecurity = $true,
        [Parameter(Mandatory=$false, Position=4,HelpMessage="Optional array of members to add (accepts same objects as Add-ADGroupMember)")] 
            [Object[]] $Members,
        [Parameter(Mandatory=$true,Position=5,HelpMessage="The OU where the permissions groups will be created")] 
            [String]$PermissionsOU = ""
	)
	Process {
        $printerName = $Name
        $printerADName = "PRN-" + $printerName
        $printerGroup = New-ADGroup -DisplayName $printerADName -GroupCategory Security -GroupScope Global -Path PermissionsOU -Name $printerADName -SamAccountName $printerADName -PassThru
        if ($Members) {
            Add-ADGroupMember -Identity $printerADName -Members $Members
        }
        $SID = $printerGroup.SID
        $printerSDLL = "G:SYD:(A;;SWRC;;;WD)(A;CIIO;RC;;;CO)(A;OIIO;RPWPSDRCWDWO;;;CO)(A;OIIO;RPWPSDRCWDWO;;;BA)(A;;LCSWSDRCWDWO;;;BA)(A;;SWRC;;;$SID)(A;CIIO;RC;;;$SID)(A;OIIO;RPWPSDRCWDWO;;;$SID)"
        Set-Printer -ComputerName $ComputerName -Name $printerName -PermissionSDDL $printerSDLL
	}
}

Export-ModuleMember -Function *