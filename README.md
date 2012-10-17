PermissionGroups
================

A collection of PowerShell modules and scripts supporting using AD groups in managing permissions to resources throughout the environment.

NOTE:  Currently a number of organization specific settings and defaults are hardcoded in the modules and WILL NEED to be edited prior to use outside of the Whatcom Transportation Authority.  Work is underway to refactor the code to either eliminate or make this eiser to change.

FolderPermissionGroups
----------------------
This makes it easy to associated folder permissions with a group and manage that assocation.  The groups are named according to the folder name and permission being set.

MailPermissionGroups
--------------------
This makes it easy to assocaite common MS Exchange processes with permission groups.  This module requires the Office365PowershellUtils module from our related project, https://github.com/whatcomtrans/Office365PowershellUtils.
