# PowerShellTools
Just a few PS scripts that I find useful - feel free to use them


ssh-copy-key.ps1
------------------------------------------------------------------------
This is a SIMPLE equivalent to the ssh-copy-id from Linux.

To copy your public ssh key to a remote Linux server via PowerShell use
the following syntax.

> ssh-copy-key -p 2222 usr@192.168.116.xx


getInstalledSoftwareList.ps1
------------------------------------------------------------------------
Creates a CSV inventory of installed software on the local Windows machine.

The script scans the standard uninstall registry locations, filters out
system-component noise by default, and writes `InstalledSoftware.csv`
beside the script.

Columns in the CSV:

> Softwarename, Current Version, Latest Version

If `winget` is installed, the script will try to populate `Latest Version`
for apps where a trustworthy upgrade match exists.

Examples:

> .\getInstalledSoftwareList.ps1

> .\getInstalledSoftwareList.ps1 -OutputPath C:\Temp\InstalledSoftware.csv

> .\getInstalledSoftwareList.ps1 -IncludeSystemComponents
