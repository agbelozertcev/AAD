# AAD Object Management Script

This PowerShell script demonstrate how to use the Microsoft Graph API to manage AAD objects.  This script allows you to add users and devices to AAD groups, create and delete groups. The UI is created using WPF(XAML).


Disclaimer

This script can retrieve information from your AAD tenant. Understand the impact of this script prior to running it. Script should be run using a non-production or "test" tenant account. 


Prerequisites

1. Install the AzureAD PowerShell module by running 'Install-Module AzureAD' or 'Install-Module AzureADPreview' from an elevated    PowerShell prompt
2. PowerShell v5.0 on Windows 10 x64 (PowerShell v4.0 is a minimum requirement for the scripts to function correctly)
3. First time usage of these script requires a Global Administrator of the Tenant to accept the permissions of the application


Getting Started

After the prerequisites are installed or met, perform the following steps to use this script:

1. Download the contents of the repository to your local Windows machine
2. Extract the files to a local folder (e.g. C:\App)
3. Run PowerShell x64 from the start menu
4. Browse to the directory (e.g. cd  C:\App)

Example Application script usage:
To use the script, from C:\App, run "cd C:\App"
Once in the folder run .\App.ps1 
