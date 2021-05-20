# enable-disable-Azure-AHUB-threaded

DESCRIPTION:

Enable or disable Azure Hybrid Benefit on specified virtual machines, using runspaces to provide concurrent processing. Changes are logged in .log and .csv files. 
Optionally enter params $AzureSubscriptionName, $AzureResourceGroup, and/or $VMName as filters.  
	
VMs to be processed can be entered in the form of a .txt file. It is recommended though not necessary that AzureSubscriptionName and 
AzureResourceGroup are specified as parameters when using a .txt input file. 
Alternatively VM data can be entered as an .xlsx or .xls file with required columns labelled SUBSCRIPTION, RESOURCE GROUP, and VM NAME, 
and optional AZURE HYBRID BENEFIT to indicate the current status of the VM. As with .txt files, parameters are available to filter the
VMs from the excel file that are processed. 

SUBSCRIPTION, RESOURCE GROUP, VM NAME - respectively contain the subscription name, resource group name, and VM name
AZURE HYBRID BENEFIT - contains the status of the VM, one of 'enabled', 'disabled', or blank 
	
	
EXAMPLES

Default state. Disable all VMs across all subscriptions, changes simulated only. 

./enable-disable-AHUB.ps1
	
Enable VMs from a .txt file, subscription and resourcegroup known: 

./enable-disable-AHUB-threaded.ps1 -Action "enable" -AzureSubscriptionName "ENMAX Corp" -AzureResourceGroup "SD_PoC" -InputFilePath "testserv.txt" -SimulateMode $False
	
Disable VMs from a .txt file, subscription and resourcegroup known, changes SIMULATED only: 

./enable-disable-AHUB-threaded.ps1 -Action "disable" -AzureSubscriptionName "ENMAX Corp" -AzureResourceGroup "SD_PoC" -InputFilePath "testserv.txt"
	
Enable VMs from an .xlsx file with the name 'SDTEST1' in simulate mode

./enable-disable-AHUB-threaded.ps1 -Action "enable" -VMName "SDTEST1" -InputFilePath "C:\Users\gzhang2\Documents\PowerShell\testserv.xlsx" 
	
Enable a specific VM in simulate mode and create .log and .csv files with unique names. 

./enable-disable-AHUB-threaded.ps1 -Action "enable" -AzureSubscriptionName "ENMAX Corp" -AzureResourceGroup "SD_PoC" -VMName "SDTEST1" -UniqueFileNames
	
Enable VMs from a .txt file, subscription and resourcegroup unknown, changes simulated only:

./enable-disable-AHUB-threaded.ps1 -Action "enable" -InputFilePath "testserv.txt"
	
Enable only VMs with the specified subscription, resource group, and name from a .txt file, changes SIMULATED only: 

./enable-disable-AHUB-threaded.ps1 -Action "enable" -AzureSubscriptionName "ENMAX Corp" -AzureResourceGroup "SD_PoC" -VMName "SDTEST1" -InputFilePath "testserv.txt"

NOTES

Dependencies: AzureRM Module v3.2 or above

Install the Azure Resource Manager modules from the PowerShell Gallery: 

Install-Module -Name AzureRM

Or if already installed, Update the Azure Resource Manager modules from the PowerShell Gallery 

Update-Module -Name AzureRM
