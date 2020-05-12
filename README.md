# AAD Object Management Script

This PowerShell script demonstrate how to use the Microsoft Graph API to manage AAD objects.  This script allows you to add users and devices to AAD groups, create and delete groups and shows some information about your tenant. The UI is created using WPF(XAML).


#### Disclaimer

This script can retrieve and change information from your AAD tenant. Understand the impact of this script prior to running it. Script should be run using a non-production or "test" tenant account. The script is provided AS IS without warranty of any kind.


## Prerequisites

* Install the AzureAD PowerShell module by running 'Install-Module AzureAD' or 'Install-Module AzureADPreview' from an elevated     PowerShell prompt
* PowerShell v5.0 on Windows 10 x64
* First time usage of this script requires a Global Administrator of the AAD tenant to accept the permissions of the application


## Getting Started

After the prerequisites are installed or met, perform the following steps to use this script:

* Download the contents of the repository to your local Windows machine
* Extract the files to a local folder (e.g. C:\App)
* Run PowerShell console from the start menu
* Browse to the directory (e.g. cd  C:\App)

Example script usage:
To use the script, from C:\App, run "cd C:\App"
Once in the folder run .\App.ps1 

## Limitations

This script uses the Microsoft Intune PowerShell application with its native permissions, for example you cannot create a user using this application because it does not have appropriate permissions (User.ReadWrite.All, Directory.ReadWrite.All). 
If you want to modify this script with the ability to create users, you should create your own application with these permissions and replace the $clientId variable in the Get-AuthToken function with the id of the application that you created.

```
$clientId = "d1ddf0e4-d672-4dae-b554-9d5bdfd93547"
```

## Additional resources

* [Microsoft Graph API documentation](https://docs.microsoft.com/en-us/graph/overview)
* [Microsoft Graph permissions reference](https://docs.microsoft.com/en-us/graph/permissions-reference)


