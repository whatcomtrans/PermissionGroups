#
# Module manifest for module 'PermissionGroups'
#
# Generated by: R. Josh Nylander
#
# Generated on: 12/3/2012
#

@{

# Script module or binary module file associated with this manifest
# RootModule = ''

# Version number of this module.
ModuleVersion = '1.2'

# ID used to uniquely identify this module
GUID = 'b37af93f-5a35-4667-b40f-1f336c662a60'

# Author of this module
Author = 'R. Josh Nylander'

# Company or vendor of this module
CompanyName = 'Whatcom Transportation Authority'

# Copyright statement for this module
Copyright = '(c) 2012 WTA. All rights reserved.'

# Description of the functionality provided by this module
Description = 'Simplifies creation of AD groups designed for manager permissions in a role groups/permissions group implementation.'

# Minimum version of the Windows PowerShell engine required by this module
PowerShellVersion = '3.0'

# Name of the Windows PowerShell host required by this module
# PowerShellHostName = ''

# Minimum version of the Windows PowerShell host required by this module
# PowerShellHostVersion = ''

# Minimum version of the .NET Framework required by this module
# DotNetFrameworkVersion = ''

# Minimum version of the common language runtime (CLR) required by this module
# CLRVersion = ''

# Processor architecture (None, X86, Amd64) required by this module
# ProcessorArchitecture = ''

# Modules that must be imported into the global environment prior to importing this module
RequiredModules = @(
    "Office365PowerShellUtils",
    "ActiveDirectory",
    "Microsoft.Online.SharePoint.PowerShell"
)

# Assemblies that must be loaded prior to importing this module
# RequiredAssemblies = @()

# Script files (.ps1) that are run in the caller's environment prior to importing this module
# ScriptsToProcess = @()

# Type files (.ps1xml) to be loaded when importing this module
# TypesToProcess = @()

# Format files (.ps1xml) to be loaded when importing this module
# FormatsToProcess = @()

# Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
NestedModules = @("MailPermissionGroups.psm1", "FolderPermissionGroups.psm1", "PermissionGroupsHelpers.psm1", "SQLReportingServices.psm1", "SharePointPermissionGroups.psm1", "PrinterPermissionGroups.psm1")

# Functions to export from this module
FunctionsToExport = '*'

# Cmdlets to export from this module
CmdletsToExport = '*'

# Variables to export from this module
VariablesToExport = '*'

# Aliases to export from this module
AliasesToExport = '*'

# Commands to export from this module as Workflows
# ExportAsWorkflow = @()

# List of all modules packaged with this module
# ModuleList = @()

# List of all files packaged with this module
# FileList = @()

# Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
PrivateData = @{

    PSData = @{

        # Tags applied to this module. These help with module discovery in online galleries.
        # Tags = @()

        # A URL to the license for this module.
        # LicenseUri = ''

        # A URL to the main website for this project.
        # ProjectUri = ''

        # A URL to an icon representing this module.
        # IconUri = ''

        # ReleaseNotes of this module
        # ReleaseNotes = ''

        ExternalModuleDependencies = @(
            "ActiveDirectory",
            "Microsoft.Online.SharePoint.PowerShell"
        )

    } # End of PSData hashtable

} # End of PrivateData hashtable
# HelpInfo URI of this module
# HelpInfoURI = ''

# Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
# DefaultCommandPrefix = ''

}

