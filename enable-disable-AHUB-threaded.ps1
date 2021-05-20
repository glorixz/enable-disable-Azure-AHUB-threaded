##########################################################################################################
<#
.SYNOPSIS
    Author: Gloria Zhang     
    Version:   
    Created: //TODO    
    Updated:    

.DESCRIPTION
	Enable or disable Azure Hybrid Benefit on specified virtual machines. Changes are logged in .log and .csv files. 
	Optionally enter params $AzureSubscriptionName, $AzureResourceGroup, and/or $VMName as filters.  
	
	VMs to be processed can be entered in the form of a .txt file. It is recommended though not necessary that AzureSubscriptionName and 
	AzureResourceGroup are specified as parameters when using a .txt input file. 
	Alternatively VM data can be entered as an .xlsx or .xls file with required columns labelled SUBSCRIPTION, RESOURCE GROUP, and VM NAME, 
	and optional AZURE HYBRID BENEFIT to indicate the current status of the VM. As with .txt files, parameters are available to filter the
	VMs from the excel file that are processed. 
	SUBSCRIPTION, RESOURCE GROUP, VM NAME - respectively contain the subscription name, resource group name, and VM name
	AZURE HYBRID BENEFIT - contains the status of the VM, one of 'enabled', 'disabled', or blank 
	
	
.EXAMPLE
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

.NOTES
    Dependencies: AzureRM Module v3.2 or above
	Install the Azure Resource Manager modules from the PowerShell Gallery: 
	>>Install-Module -Name AzureRM
	Or if already installed, Update the Azure Resource Manager modules from the PowerShell Gallery 
	>>Update-Module -Name AzureRM
	
	Scripts referenced: 
	Azure - Enable Hybrid Use Benefit On All VMs in All Subscriptions by Neil Bird, MSFT
	Enable-Disable Azure Hybrid Benefit(AHUB) for Windows VM and Windows VMSS by Pradebban Raja
	
#>
##########################################################################################################

[CmdletBinding()]
# define script parameters 
Param(
	# Azure subscription. Leave blank to operate in all or multiple subscriptions. 
	[Parameter(Position=0,Mandatory=$false)][String]$AzureSubscriptionName="null",
	# Resource group name. Specify the resource group to operate on.
	[Parameter(Position=1,Mandatory=$false)][String]$AzureResourceGroup="null", 
	# VM name. Specify the VM to operate on.
	[Parameter(Position=2,Mandatory=$false)][String]$VMName="null",
	# Specify whether to enable or disable AHUB. Default: disable
	[Parameter(Position=4,Mandatory=$false)][String]$Action="disable",
	# Include to make changes rather than simply generate a report of changes that would be made. 
	[Parameter(Position=5,Mandatory=$false)][bool]$SimulateMode=$True,	
	# Folder Path for Output, if not specified defaults to script folder. e.g. C:\Scripts\
	[parameter(Position=6,Mandatory=$false)][string]$OutputFolderPath="null",  
	# enable unique file names for CSV files optionally 
	[parameter(Position=7,Mandatory=$false)][switch]$UniqueFileNames,
	# Folder Path for input file. 
	[parameter(Position=8,Mandatory=$false)][string]$InputFilePath="null"  
)

# Set strict mode to identify typographical errors and facilitate troubleshooting
Set-StrictMode -Version Latest

# Set Error and Warning action preferences
$ErrorActionPreference = "Stop"
$WarningPreference = "Stop"

##########################################################################################################

#######################################
## FUNCTION - Set-OutputLogFiles  
#######################################
# Generate unique log file names. Sets $OutputFolderFilePath and $OutputFolderFilePathCSV. 
#######################################
Function Set-OutputLogFiles {

    [string]$FileNameDataTime = Get-Date -Format "yy-MM-dd_HHmmss"
            
    ### default to script folder, or user profile folder.
    if([string]::IsNullOrWhiteSpace($script:MyInvocation.MyCommand.Path)){
        $ScriptDir = "."
    } else {
        $ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path
    }

	### set the $OutputFolderFilePath variable
    if(Test-Path($OutputFolderPath)){ # param has been set and is valid
        if($UniqueFileNames) {
            $script:OutputFolderFilePath = "$OutputFolderPath\enable-disable-AHUB_$($FileNameDataTime).log"
        } else {
            $script:OutputFolderFilePath = "$OutputFolderPath\enable-disable-AHUB.log"
        }
    } else { # Folder specified is not valid, default to script or user profile folder.
        $OutputFolderPath = $ScriptDir
        if($UniqueFileNames) {
            $script:OutputFolderFilePath = "$($ScriptDir)\enable-disable-AHUB_$($FileNameDataTime).log"
        } else {
            $script:OutputFolderFilePath = "$($ScriptDir)\enable-disable-AHUB.log"
        }
    }

    ### set the $OutputFolderFilePathCSV variable, making it unique if $CSVUniqueFileNames is specified
    if($UniqueFileNames) {
        $script:OutputFolderFilePathCSV = "$OutputFolderPath\enable-disable-AHUB-audit-report_$($FileNameDataTime).csv"
    } else {
        $script:OutputFolderFilePathCSV = "$OutputFolderPath\enable-disable-AHUB-audit-report.csv"
    }
}


###################################
## FUNCTION - Out-ToHostAndFile 
###################################
# Used to create a transcript of output. 
# Requires that Set-OutputLogFiles is called first.  
###################################
Function Out-ToHostAndFile {
    Param(
	    [parameter(Position=0,Mandatory=$True)][string]$Content, # the content to be written
        [parameter(Position=1,Mandatory=$False)][string]$FontColour="White",
        [parameter(Position=2)][switch]$NoNewLine # specify to write without starting a new line 
    )

    ### Write Content to Output File
    try {
		if($NoNewLine) {  # write with the NoNewLine switch if specified
			Out-File -FilePath $OutputFolderFilePath -Encoding UTF8 -Append -InputObject $Content -NoNewline -ErrorAction $ErrorActionPreference
		} else {
			Out-File -FilePath $OutputFolderFilePath -Encoding UTF8 -Append -InputObject $Content -ErrorAction $ErrorActionPreference
		}           
	} catch [System.Management.Automation.CmdletInvocationException] {
		# being used by another process. ---> System.IO.IOException
		# timing issue with locked file, attempt to write again
		Start-Sleep -Milliseconds 250
		Out-File -FilePath $OutputFolderFilePath -Encoding UTF8 -Append -InputObject $Content -NoNewline -ErrorAction $ErrorActionPreference
	}
	
	### Write Content to the host
    if($NoNewLine) {
        Write-Host $Content -ForegroundColor $FontColour -NoNewline
    } else {
        Write-Host $Content -ForegroundColor $FontColour
    }
}


###################################
## FUNCTION - Get-AzureConnection 
###################################
# Set the AzureRmContext with a subscription. Run once as part of the initialization process. 
# Error handling: user has not logged in, subscription is invalid.
# By default, selects the first subscription found (use case: VMs in multiple subscriptions).
###################################
Function Get-AzureConnection{
    ### Set error-handling messages 
    if ($SimulateMode) {
        $GridViewTile = "Select the subscription to Simulate Enabling Azure Hybrid Use Benefit"
    } else {
        $GridViewTile = "Select the subscription to Enable Azure Hybrid Use Benefit" 
    }

	### Ensure that the user is logged in to an Azure session. 
	Try{ # check for an active session
		$test = Get-AzureRmsubscription -ErrorAction Continue
	} Catch [System.Management.Automation.PSInvalidOperationException]{
		# if not logged in, do so. 
		Login-AzureRmAccount -ErrorAction $ErrorActionPreference
	}
	
	### log successful login
    Out-ToHostAndFile "SUCCESS: " "Green" -nonewline; `
    Out-ToHostAndFile "Logged into Azure using Account ID: " -NoNewline; `
    Out-ToHostAndFile (Get-AzureRmContext).Account.Id "Green"
    Out-ToHostAndFile " "
}

###################################
## FUNCTION - Process-txt
###################################
# Convert the lines of an input txt file to array form. Sets the $machines variable.
# Error handling: invalid file path 
# Format of $machines: array1<array2<string>> where array2=(subscription, resource group, VM name)
###################################
Function Process-txt {
    ### create array and convert it to a collection so that remove function is available
	$temp = @()

    ### ensure that the input file path is valid
    while (-Not (Test-Path($InputFilePath) -ErrorAction SilentlyContinue)){ # if input file path is not valid
        # prompt the user for a correct path 
		Out-ToHostAndFile "$InputFilePath is invalid. Please enter a valid file path, or type 'exit' to end the script." "Red"
		Out-ToHostAndFile " "
		$FilePathResponse = Read-Host -Prompt "InputFilePath:`n`n[<filepath>]       [exit]"
		if ($FilePathResponse.ToLower() -eq "exit"){
			Out-ToHostAndFile " "
			Out-ToHostAndFile "User typed 'exit'. Exiting script..."
			Exit
		} else {
			$Script:InputFilePath = $FilePathResponse
			Out-ToHostAndFile "`nProcessing input file: $($InputFilePath)" "Green"
		}
    }

    ### append each machine name to $temp as part of a sub-array with subscription name and resource group. 
    Get-Content -Path $InputFilePath | ForEach-Object {
        $tempArray = "", "", $_.Trim()
        $temp += ,$tempArray
    }
	$script:machines = {$temp}.Invoke()
}

###################################
## FUNCTION - Process-xl
###################################
# Convert the rows of an input xlsx or xls file to array form. Sets the $machines variable.
# Error handling: invalid file path, missing columns 
# Format of $machines: array1<array2<string>> where array2=(subscription, resource group, VM name)
###################################
Function Process-xl {
    #declare empty array
    $temp = @()
	
	### ensure that the input file path is valid
    while (-Not (Test-Path($InputFilePath) -ErrorAction SilentlyContinue)){ # if input file path is not valid
        # prompt the user for a correct path 
		Out-ToHostAndFile "$InputFilePath is invalid. Please enter a valid file path, or type 'exit' to end the script." "Red"
		Out-ToHostAndFile "Note: If using excel input, a full file path is required, e.g. 'C:\Users\bobross\Documents\myfile.xlsx'"
		Out-ToHostAndFile " "
		$FilePathResponse = Read-Host -Prompt "InputFilePath:`n`n[<filepath>]       [exit]"
		if ($FilePathResponse.ToLower() -eq "exit"){
			Out-ToHostAndFile " "
			Out-ToHostAndFile "User typed 'exit'. Exiting script..."
			Exit
		} else {
			$Script:InputFilePath = $FilePathResponse
			Out-ToHostAndFile "`nProcessing input file: $($InputFilePath)" "Green"
		}
    }
	
    ### initialization
    $objExcel = New-Object -ComObject Excel.Application # Create an Object Excel.Application using Com interface
	$objExcel.Visible = $false  # Disable the 'visible' property so the document won't open in excel
	while ($True){ # ensure that a full file path has been entered
        try {
            $WorkBook = $objExcel.Workbooks.Open($InputFilePath) # Open the Excel file
			break
        } catch {
            if($error[0].Exception.ToString().Contains("Sorry, we couldn't find")) { # if file path can't be used               
                Out-ToHostAndFile "$InputFilePath is invalid. Please enter a valid file path, or type 'exit' to end the script." "Red"
				Out-ToHostAndFile "Note: If using excel input, a full file path is required, e.g. 'C:\Users\bobross\Documents\myfile.xlsx'"
				Out-ToHostAndFile " " 
				# prompt user to enter a valid file path
                $FilePathResponse = Read-Host -Prompt "InputFilePath:`n`n[<filepath>]       [exit]"
                if ($FilePathResponse.ToLower() -eq "exit"){
			        Out-ToHostAndFile " "
			        Out-ToHostAndFile "User typed 'exit'. Exiting script..."
			        Exit
	            } else {
		            $Script:InputFilePath = $FilePathResponse
		            Out-ToHostAndFile "`nProcessing input file: $($InputFilePath)"
	            } 
            } else { # All other errors.
                Out-ToHostAndFile "Error: $($error[0].Exception)"
                Exit  
            }
        }
    }
	$sheet = $WorkBook.Sheets.Item(1) # create a variable for the first sheet in the file 
	
	### find and save the cells of the SUBSCRIPTION, RESOURCE GROUP, VM NAME, and AZURE HYBRID BENEFIT columns
	$SUBSCRIPTION = $sheet.Cells.Find('SUBSCRIPTION')
	$RESOURCE_GROUP = $sheet.Cells.Find('RESOURCE GROUP')
	$VM_NAME = $sheet.Cells.Find('VM NAME')
	$AZURE_HYBRID_BENEFIT = $sheet.Cells.Find('CURRENT AZURE HYBRID BENEFIT')
	
	# exit if SUBSCRIPTION, RESOURCE GROUP or VM NAME column is missing
	if (-Not $VM_NAME -Or -Not $RESOURCE_GROUP -Or -Not $SUBSCRIPTION) {
		Out-ToHostAndFile "Error: Excel file does not contain one of the columns 'VM NAME', 'RESOURCE GROUP', or 'SUBSCRIPTION'. Exiting script..." "Red"
		Exit
	}
	
	# initialize loop variables
	$header_row = $VM_NAME.row 
	$objRange = $sheet.UsedRange
	$lastRow = $objRange.SpecialCells(11).row  # get the index of the last row. '11' is the XlCellType that gets the last cell in objRange.  
	
	### loop through rows, adding VMs to array as necessary
	for ($rowidx = $header_row +1; $rowidx -le $lastRow; $rowidx++){
		if ($AZURE_HYBRID_BENEFIT){ # if the AZURE HYBRID BENEFIT column is present
			# retrieve its value 
			[string]$vm_ahub_val = $sheet.Cells.Item($rowidx, $AZURE_HYBRID_BENEFIT.column).text
			if ($vm_ahub_val.ToLower() -eq ($Action.ToLower() + "d")){ # if AHUB is recorded as having the target status already
				# skip this row
				Continue
			}
		}
		
		# retrieve subscription, resource group, and vm name values. 
		[string]$vm_name_val = $sheet.Cells.Item($rowidx, $VM_NAME.column).text
		[string]$vm_resg_val = $sheet.Cells.Item($rowidx, $RESOURCE_GROUP.column).text
		[string]$vm_subs_val = $sheet.Cells.Item($rowidx, $SUBSCRIPTION.column).text
		
		# handle null VM name
		if (-Not $vm_name_val) {
			# skip this row if the VM NAME column has no associated value
			Continue
		} 
		
		# else, add the current row to the array. 
		$tempArray = $vm_subs_val, $vm_resg_val, $vm_name_val
		$temp += ,$tempArray
	} # end loop
	
	# convert array to a collection so that remove function is available
	$script:machines = {$temp}.Invoke()
	
	# close the workbook
	$objExcel.Workbooks.Close()
}


################################### 
## SCRIPTBLOCK - ProcessVM
###################################
#  Given a PSVirtualMachine, modify its LicenceType. Returns updates to $Script:SuccessCount, $Script:FailedCount, $Script:VMNotCompatibleCount, and $Script:AlreadyEnabledOrDisabledCount. 
###################################
$ProcessVM = {
	param (
        [Parameter(Mandatory=$true)]$RmVM,
		[Parameter(Mandatory=$true)][string]$ResourceGroup,
		[Parameter(Mandatory=$true)]$Subscription,
		
		[Parameter(Mandatory=$true)][string]$TargetLicenceType,
		[Parameter(Mandatory=$true)][bool]$SimulateMode, 
		[Parameter(Mandatory=$true)][string]$Action,  
		[Parameter(Mandatory=$true)][string]$OutputFolderFilePathCSV  
    )
	
	# Create New Ordered Hash Table to store VM details
	$VMHUBOutput = [ordered]@{}
	$VMHUBOutput.Add("Resource Group",$ResourceGroup)
	$VMHUBOutput.Add("VM Name",$RmVM.Name)
	$VMHUBOutput.Add("VM Size",$RmVM.HardwareProfile.VmSize)
	$VMHUBOutput.Add("VM Location",$RmVM.Location)
	$VMHUBOutput.Add("OS Type",$RmVM.StorageProfile.OsDisk.OsType)
	
	# Create an array to hold messages generated during this function 
	$resultmsgs = @()
	
	# initialize counts that will be returned 
	$SuccessCount = 0
	$FailedCount = 0
	$VMNotCompatibleCount = 0
	$AlreadyEnabledOrDisabledCount = 0

	# If the VM is a Windows VM
	if($RmVM.StorageProfile.OsDisk.OsType -eq "Windows") {
		#write-host "Windows OS confirmed"
		# If HUB is NOT already enabled/disabled
		if($RmVM.LicenseType -ne $TargetLicenceType -And $RmVM.LicenseType -ne "Windows_Client") {
			#write-host "not already toggled confirmed"
			if($SimulateMode) { # $SimulateMode set to $True (default), No Updates will be performed
				$resultmsgs += ,("`tINFO: ", "Green", "-NoNewLine")
				$resultmsgs += ,("Would $($Action) HUB on VM: $($RmVM.Name)", $null, $null)
				if ($Action.ToLower() -eq "enable"){
					$VMHUBOutput.Add("HUB Enabled","No")
				} elseif ($Action.ToLower() -eq "disable"){
					$VMHUBOutput.Add("HUB Enabled","Yes")
				}
				$VMHUBOutput.Add("Script Action","Script would $($Action) HUB")
				# Increment counter, for reporting only
				$SuccessCount++
			} else { # $SimulateMode set to $False, updates will be performed
				$resultmsgs += ,("`tUpdating $($RmVM.Name)...", $null, $null)
				
				$RmVM.LicenseType = $TargetLicenceType
				$AzureHUB = (Update-AzureRmVM -ResourceGroupName $ResourceGroup -VM $RmVM -ErrorVariable UpdateVMFailed -ErrorAction SilentlyContinue)
				
				if($UpdateVMFailed) { # Failed to toggle HUB, unhandled error
					$FailedCount++
					$resultmsgs += ,("`tERROR: ", "Red", "-NoNewLine")
					$resultmsgs += ,("`t$($RmVM.Name) - Failed to set LicenseType...", $null, $null)
					$resultmsgs += ,("`tError: $($UpdateVMFailed.Exception)", $null, $null)
					$VMHUBOutput.Add("HUB Enabled","No")
					$VMHUBOutput.Add("Script Action","Failed to set LicenseType: $($UpdateVMFailed.Exception)")
				} else { # LicenceType change successful
					if($AzureHUB.IsSuccessStatusCode -eq $True) { # Successfully enabled HUB
						$SuccessCount++
						$resultmsgs += ,("`tSUCCESS: ", "Green", "-NoNewLine")
						$resultmsgs += ,("$($RmVM.Name) LicenseType set to $($Action) HUB", $null, $null)
						if ($Action.ToLower() -eq "enable"){
							$VMHUBOutput.Add("HUB Enabled","Yes")
						} elseif ($Action.ToLower() -eq "disable"){
							$VMHUBOutput.Add("HUB Enabled","No")
						}
						$VMHUBOutput.Add("Script Action","HUB $($Action)d Successfully")
					
					} elseif($AzureHUB.StatusCode.value__ -eq 409) {
						$VMNotCompatibleCount++
						# Marketplace VM Image with additional software, such as SQL Server
						$resultmsgs += ,("`tINFO: ", "Yellow", "-NoNewLine")
						$resultmsgs += ,("$($RmVM.Name) is NOT compatible with Azure HUB", $null, $null)
						$VMHUBOutput.Add("HUB Enabled","No")
						$VMHUBOutput.Add("Script Action","Marketplace VM, NOT compatible with Azure HUB")
					} else { # Failed to enabled HUB, unhandled error
						$FailedCount++
						$resultmsgs += ,("`tERROR: ", "Red", "-NoNewLine")
						$resultmsgs += ,("`t$($RmVM.Name) - Failed to set LicenseType...", $null, $null)
						$resultmsgs += ,("`tStatusCode = $AzureHUB.StatusCode ReasonPhrase = $AzureHUB.ReasonPhrase", $null, $null)
						if ($Action.ToLower() -eq "enable"){
							$VMHUBOutput.Add("HUB Enabled","No")
						} elseif ($Action.ToLower() -eq "disable"){
							$VMHUBOutput.Add("HUB Enabled","Yes")
						}
						$VMHUBOutput.Add("Script Action","Failed to set LicenseType: $($AzureHUB.StatusCode)")
					} #end if
				} #end if
			} #end if
		} else { # the target HUB LicenceType has already been set
			$AlreadyEnabledOrDisabledCount++
			$resultmsgs += ,("`tINFO: ", "Yellow", "-NoNewLine")
			$resultmsgs += ,("$($RmVM.Name) already has HUB LicenseType $($Action)d", $null, $null)
			if ($Action.ToLower() -eq "enable"){
				$VMHUBOutput.Add("HUB Enabled","Yes")
			} elseif ($Action.ToLower() -eq "disable"){
				$VMHUBOutput.Add("HUB Enabled","No")
			}
			$VMHUBOutput.Add("Script Action","HUB LicenseType Already $($Action)d")
		} #end if
	} elseif($RmVM.StorageProfile.OsDisk.OsType -eq "Linux") { # OS is Linux instead of Windows                          
		$VMNotCompatibleCount++
		$resultmsgs += ,("`tINFO: ", "Yellow", "-NoNewLine")
		$resultmsgs += ,("$($RmVM.Name) is running a Linux OS", $null, $null)
		$VMHUBOutput.Add("HUB Enabled","NA")
		$VMHUBOutput.Add("Script Action","Linux VM, NOT compatible with Azure HUB")
	} else { # Non-Windows / Non-Linux VM
		$VMNotCompatibleCount++
		$resultmsgs += ,("`tINFO: ", "Yellow", "-NoNewLine")
		$resultmsgs += ,("$($RmVM.Name) is NOT running a $($RmVM.StorageProfile.OsDisk.OsType) OS", $null, $null)
		$VMHUBOutput.Add("HUB Enabled","NA")
		$VMHUBOutput.Add("Script Action","Non-Windows VM, NOT compatible with Azure HUB")
	} # end if

	# Add Subscription Name and Subscription ID
	$VMHUBOutput.Add("Subscription Name",$Subscription.Name)
	$VMHUBOutput.Add("Subscription ID",$Subscription.Id)

	### Export VM to CSV File
	
	$Data = @() # Create an empty Array to hold Hash Table
	$Row = New-Object PSObject # Create a PSObject to store data in key-value form 
	$VMHUBOutput.GetEnumerator() | ForEach-Object {# Loop Hash Table and add to PSObject
		$Row | Add-Member NoteProperty -Name $_.Name -Value $_.Value
    }

	# cast PSObject to Array
	$Data = $Row

	# Export Array to CSV
    try {
        $Data | Export-CSV -Path $OutputFolderFilePathCSV -Encoding UTF8 -NoTypeInformation -Append -Force -ErrorAction $ErrorActionPreference
    } catch {
        # On first error, attempt to write again
        $Data | Export-CSV -Path $OutputFolderFilePathCSV -Encoding UTF8 -NoTypeInformation -Append -Force -ErrorAction $ErrorActionPreference
    }
	
	### create return object 
    $RunResult = New-Object PSObject -Property @{
        Messages = $resultmsgs
		SuccessCount = $SuccessCount
		FailedCount = $FailedCount
		VMNotCompatibleCount = $VMNotCompatibleCount 
		AlreadyEnabledOrDisabledCount = $AlreadyEnabledOrDisabledCount
    }
    return $RunResult
}


################################### 
## FUNCTION - ProcessSubscription
###################################
#  Given a subscription, process the resource groups within it. 
###################################
Function ProcessSubscription {
	param (
		[Parameter(Mandatory=$true)][string]$AzureSubscriptionName
		#[Parameter(Mandatory=$true)][string]$machines (script-level)
		#[Parameter(Mandatory=$true)][string]$AzureResourceGroup (script-level)
		#[Parameter(Mandatory=$true)][string]$ErrorActionPreference (script-level)
    )
	
	try {
		$AzureSubscription = Get-AzureRmSubscription -SubscriptionName $AzureSubscriptionName -ErrorAction $ErrorActionPreference
	} Catch {
		if($error[0].Exception.ToString().Contains("was not found")){
			Out-ToHostAndFile "Subscription does not exist. Continuing..."
			Continue
		} else {
			Write-Error "Error: $($error[0].Exception)"
			Exit	
		}
	}

	### Set AzureRMContext for the current subscription
	Set-AzureRmContext -SubscriptionId $AzureSubscription.Id
	
	$SubscriptionMachines = @()
	### filter out VMs from $machines that are in the current subscription if an input file was used. 
	if ($machines){ # if input file was used	
		ForEach ($machine in $machines) {
			# include all machines that are or might be in the current subscription.
			# Each subscription gets a list of all the null-subscription machines
			#    in addition to the input machines that are specifically in that subscription. 
			if ($machine[0] -eq "" -OR $machine[0] -eq $AzureSubscription.Name){
				$SubscriptionMachines += ,$machine
			}
		}
		if (-Not $SubscriptionMachines){ #check that this subscription contains machines to process. 
			# if not, continue to the next subscription.
			Out-ToHostAndFile "There are no VMs from the input file to process in this subscription. Continuing..."				
			Continue
		}
	}
	
	### populate a list of resource group names for the current subscription 
	if ($AzureResourceGroup -eq "null") { # if no resource group specified
		# get the names of all resource groups in this subscription 
		[array]$ResourceGroups = (Get-AzureRmResourceGroup -ErrorAction $ErrorActionPreference).ResourceGroupName
	} else { # if resource group specified
		try { # try to get the resource group with the specified name
			[array]$ResourceGroups = (Get-AzureRmResourceGroup -name $AzureResourceGroup -ErrorAction $ErrorActionPreference).ResourceGroupName
		} Catch {
			if($error[0].Exception.ToString().Contains("Provided resource group does not exist")) {                   
				# if subscription does not contain a resource group of the specified name, skip this subscription.
				Out-ToHostAndFile "Subscription does not contain the specified resource group. Continuing..."					
				Continue  
			} else { #unhandled errors.
				Write-Error "Error: $($error[0].Exception)"
				Exit		
			}
		}		    
	}
	
	# check that the subscription contains at least one resource group
	if(-Not $ResourceGroups) { 
		#if not, continue to next subscription.
		Out-ToHostAndFile "There are no resource groups to process in this subscription. Continuing..."			
		Continue
	}
	
	
	# create array to hold messages and jobs created while processing resource groups. Structure: <<messages, jobs>> where messages=<<string>> and jobs=<<job objects>>
	$RGoutput = @()
	
	### initialize runspace pool for VM processing
	$VMRunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $maxVMThreads)
	$VMRunspacePool.Open()
	
	Out-ToHostAndFile " "
	Out-ToHostAndFile "Processing Resource Groups in ""$($AzureSubscription.Name)""..."
	$ResourceGroupCounter = 0
	
	### Loop through $ResourceGroups
	ForEach($ResourceGroup in $ResourceGroups) {
	
		# progress message
		$ResourceGroupCounter++
		
		Out-ToHostAndFile "Creating jobs for resource group: " -NoNewLine
		Out-ToHostAndFile "$ResourceGroupCounter of $($ResourceGroups.count)" "Green"
		Out-ToHostAndFile "Group name: " -NoNewLine
		Out-ToHostAndFile "$ResourceGroup" "Yellow"
		Out-ToHostAndFile " "
		
		# Create an array to hold messages generated during this function 
		$resultmsgs = @()
		
		$ResourceMachines = @() #array of INPUT machines relevant to this resource
		### if input file specified, populate a list of machines ($ResourceMachines) that are or may be in the current resource group
		# Each resource gets a list of all machines with null resource group 
		#    in addition to the subscription's input machines that are specifically in that resource group.
		if ($SubscriptionMachines){
			foreach ($machine in $SubscriptionMachines){
				# include all machines that are or might be in the current resource group. 
				if ($machine[1] -eq "" -OR $machine[1] -eq $ResourceGroup){
					$ResourceMachines+= ,$machine
				}
			}
			if (-Not $ResourceMachines){ #check that this resource contains machines to process. 
				# if not, continue to the next resource.
				$resultmsgs += ,("There are no VMs to process in resource group '$ResourceGroup'. Continuing...", $null, $null)	
				$output = $resultmsgs, $null
				$RGoutput += ,$output					
				Continue
			}
		}
		
		[array]$RmVMs = @() # array of machines in this resource that will be processed
		
		# Get virtual machines in the given Resource Group
		if ($VMName -eq "null") { # if no VM specified
			if ($ResourceMachines){ # input file specified
				[array]$VMsInThisGroup = (Get-AzureRmVM -ResourceGroupName $ResourceGroup -ErrorAction $ErrorActionPreference).Name
				# Check every candidate machine to see if it is in the current Resource Group
				ForEach ($machine in $ResourceMachines) {
					if($VMsInThisGroup.contains($machine[2])){# if resource group contains the machine
						try { # add the VM to $RmVMs
							$RmVMs += Get-AzureRmVM -ResourceGroupName $ResourceGroup -Name $machine[2] -ErrorAction $ErrorActionPreference -WarningAction $WarningPreference
						} Catch {	
							Write-Error "Error: $($error[0].Exception)"
							Exit
						}
					}
				}
			} else { # if no input file specified
				# get all VMs in the resource group 
				$RmVMs = Get-AzureRmVM -ResourceGroupName $ResourceGroup -ErrorAction $ErrorActionPreference -WarningAction $WarningPreference
			}
		} else { # if VMName specified
			if ($ResourceMachines){ # input file also specified
				[array]$VMsInThisGroup = (Get-AzureRmVM -ResourceGroupName $ResourceGroup -ErrorAction $ErrorActionPreference).Name
				# Check every candidate machine to see if it is in the current Resource Group
				ForEach ($machine in $ResourceMachines) {
					if($machine[2] -eq $VMName -And $VMsInThisGroup.contains($machine[2])){# if resource group contains the machine AND has the input name $VMName
						try { # add the VM to $RmVMs
							$RmVMs += Get-AzureRmVM -ResourceGroupName $ResourceGroup -Name $machine[2] -ErrorAction $ErrorActionPreference -WarningAction $WarningPreference
						} Catch {	
							Write-Error "Error: $($error[0].Exception)"
							Exit
						}
					}
				}
			} else { #input file not specified
				try { # get the specified VM
					$RmVMs = Get-AzureRmVM -ResourceGroupName $ResourceGroup -Name $VMName
				} Catch {
					if($error[0].Exception.ToString().Contains("was not found")) {                   
						# if subscription does not contain a resource group of the specified name, skip this subscription.
						$resultmsgs += ,("Resource Group '$ResourceGroup' does not contain the specified VM. Continuing...", $null, $null)
						$output = $resultmsgs, $null
						$RGoutput += ,$output						
						Continue  
					} else { 		
						Write-Error "Error: $($error[0].Exception)"
						Exit		
					}
				}
			}
		}

		if(-Not $RmVMs) { # check that the list to process contains at least one VM
			$resultmsgs += ,("`n`tResource Group '$ResourceGroup' contains no target VMs. Continuing...", $null, $null)
			$output = $resultmsgs, $null
			$RGoutput += ,$output
			Continue
		} 
		
		$VMJobs = @() # array to hold the jobs created for this resource group
		
		### Loop through VMs
		ForEach($RmVM in $RmVMs) {
			# create a runspace for the VM
			$Job = [powershell]::Create().AddScript($ProcessVM).AddArgument($RmVM).AddArgument($ResourceGroup). `
			AddArgument($AzureSubscription).AddArgument($TargetLicenceType).AddArgument($SimulateMode). `
			AddArgument($Action).AddArgument($OutputFolderFilePathCSV)
			$Job.RunspacePool = $VMRunspacePool
			# add it to the list of VM jobs along with variables for meta-properties
			$VMJobs += New-Object PSObject -Property @{
			  Pipe = $Job
			  Result = $Job.BeginInvoke() # starts the job and stores its return value
			}
		} #end forEach VM
		
		$output = $resultmsgs, $VMJobs
	
		# append messages and jobs of this resource group to the master list
		$RGoutput += ,$output
		
	} #end forEach resource group 
	
	# for each resource group, wait until all the VMs in the resource group have finished and process the messages. 
	forEach ($rg in $RGoutput){
		# check that the resource group's jobs have been completed; otherwise, wait
		While ($rg[1] -ne $null -And $rg[1].Result.IsCompleted -contains $false){ 
			Start-Sleep -Seconds 1
		}

		# harvest messages and counts from the jobs
		ForEach ($Job in $rg[1]){
			$vmresult = $Job.Pipe.EndInvoke($Job.Result)
			
			$rg[0] += $vmresult.Messages #add vm messages to the rg messages
			
			$Script:SuccessCount += $vmresult.SuccessCount
			$Script:FailedCount += $vmresult.FailedCount
			$Script:VMNotCompatibleCount += $vmresult.VMNotCompatibleCount 
			$Script:AlreadyEnabledOrDisabledCount += $vmresult.AlreadyEnabledOrDisabledCount
			
			$Job.Pipe.Dispose()
		}
		
		#display messages
		ForEach ($msg in $rg[0]){
			$arg1 = $msg[0]
			$arg2 = $msg[1]
			$arg3 = $msg[2]

			if ($arg2 -And $arg3){
				Out-ToHostAndFile $arg1 $arg2 -NoNewLine
			} elseif ($arg2){
				Out-ToHostAndFile $arg1 $arg2 
			} elseif ($arg3){
				Out-ToHostAndFile $arg1 -NoNewLine
			} else {
				Out-ToHostAndFile $arg1
			}
		}
	
	} # end forEach rg
	
	# Dispose of runspaces
	$VMRunspacePool.Close()
}


################################### 
## FUNCTION - Run-AHUBToggleProcess
###################################
# Enable or disable azure hybrid use benefit on specified VMs. Disables by default. 
# Sets the $TargetLicenceType, $SuccessCount, $FailedCount, $AlreadyEnabledOrDisabledCount, and $VMNotCompatibleCount variables.  
###################################
Function Run-AHUBToggleProcess{
    ### Set up counters for process results
    [int]$Script:SuccessCount = 0
    [int]$Script:FailedCount = 0
    [int]$Script:AlreadyEnabledOrDisabledCount = 0
    [int]$Script:VMNotCompatibleCount = 0
	
	### set $TargetLicenceType using the $Action parameter.
	# note that LicenceType will be null for VMs that have NEVER had AHUB enabled; such VMs will read as neither enabled nor disabled
	#     (though functionally they are disabled) and will always be treated by this implementation as targets for LicenceType change. 
	while ($Action.ToLower() -ne "enable" -And $Action.ToLower() -ne "disable") {
		Out-ToHostAndFile "Action not valid. Please enter again." "Red"
        Out-ToHostAndFile " "
		$ActionResponse = Read-Host -Prompt "Action? `n`n[enable]   [disable]"
		$Script:Action = $ActionResponse
	}
	
	if ($Action.ToLower() -eq "enable"){
		[string]$Script:TargetLicenceType = "Windows_Server"
	} elseif ($Action.ToLower() -eq "disable"){
		[string]$Script:TargetLicenceType = "None"
	}
	
    ### populate $AzureSubscriptions with the target subscription(s)
    if($AzureSubscriptionName -eq "null") { # If $AzureSubscriptionName parameter has not been specified
        # Get all Subscriptions
        [array]$AzureSubscriptions = (Get-AzureRmSubscription -ErrorAction $ErrorActionPreference).Name
    } else {
        # Use the subscription that has been specified
        [array]$AzureSubscriptions += $AzureSubscriptionName
    }
	
	if (-Not $AzureSubscriptions){ # if no subscriptions are found
		Out-ToHostAndFile "Error: this Azure account has no subscriptions. Exiting script..." "Red"
		Exit
	}
	
	$script:machines = @() 
    ### process input file, if specified. Populate the $machines variable with sub-arrays containing subscription name, resource group, and VM name. 
    if ($InputFilePath -ne "null"){
		Out-ToHostAndFile "`nProcessing input file: $($InputFilePath)" "Green"
        if ([IO.Path]::GetExtension($InputFilePath) -eq ".txt"){
            Process-txt # Process a list of VMs in a txt file
        } elseif ([IO.Path]::GetExtension($InputFilePath) -eq ".xls" -Or [IO.Path]::GetExtension($InputFilePath) -eq ".xlsx"){
            Process-xl # Process a list of VMs in an excel file
        }
		if (-Not $machines){ # if $machines is empty after processing is finished
			Out-ToHostAndFile "No VMs found in input file. Exiting script..." "Red"
            Exit
		}
    }
	
	Out-ToHostAndFile " "
	
    ### Confirm that user wants to continue
    if($SimulateMode) {
        # Simulate Mode True
        Out-ToHostAndFile "INFO: " "Yellow" -NoNewLine
        Out-ToHostAndFile "Simulate Mode Enabled" "Green" -NoNewLine
        Out-ToHostAndFile " - No updates will be performed."
		Out-ToHostAndFile " Selected action: $($Action)"
        Out-ToHostAndFile " "
    } else {
        # Simulate Mode False 
        Out-ToHostAndFile "INFO: " "Yellow" -NoNewLine
        Out-ToHostAndFile "Simulate Mode DISABLED - Updates will be performed." "Green"
		Out-ToHostAndFile " Selected action: $($Action)"
        Out-ToHostAndFile " "
    }
    Do {
        $UserConfirmation = Read-Host -Prompt "Continue? `n`n[yes]   [no]"
    } While ($UserConfirmation.ToLower() -ne 'yes' -And $UserConfirmation.ToLower() -ne 'no')
    
    if($UserConfirmation.ToLower() -eq 'no'){
        Out-ToHostAndFile "`nUser typed 'no' when asked to confirm, exiting script..."
        Out-ToHostAndFile " "
        Exit
    } else {
        Out-ToHostAndFile "`nUser typed 'yes' to confirm..."
        Out-ToHostAndFile " "
    }
   
    ### loop through $AzureSubscriptions
    $SubscriptionCount = 0
    ForEach($AzureSubscription in $AzureSubscriptions) {
        $SubscriptionCount++

        Out-ToHostAndFile "`nProcessing Azure Subscription: " -NoNewLine
        Out-ToHostAndFile "$SubscriptionCount of $($AzureSubscriptions.Count)" "Green"
        Out-ToHostAndFile "Subscription Name = " "Cyan" -NoNewLine
        Out-ToHostAndFile """$($AzureSubscription)""`n" "Yellow"
        
		ProcessSubscription -AzureSubscriptionName $AzureSubscription
        #& $ProcessSubscription -AzureSubscription $AzureSubscription
    } 

	
	### Add up all of the counters and report status of script
    [int]$TotalVMsProcessed = $Script:SuccessCount + $Script:FailedCount + $Script:AlreadyEnabledOrDisabledCount `
    + $Script:VMNotCompatibleCount
    
	# Output Extension Installation Results
    Out-ToHostAndFile " "
    Out-ToHostAndFile "====================================================================="
	if ($Action -eq "enable"){ # format text according to action
		Out-ToHostAndFile "`tEnable Azure HUB LicenseType Results`n" "Green"
	} else {
		Out-ToHostAndFile "`tDisable Azure HUB LicenseType Results`n" "Green"
	}
    if($SimulateMode) { 
		if ($Action -eq "enable"){ # format text according to action
			Out-ToHostAndFile "Would have HUB enabled:`t`t`t$($Script:SuccessCount)"
		} else {
			Out-ToHostAndFile "Would have HUB disabled:`t`t$($Script:SuccessCount)"
		}
    } else {
		if ($Action -eq "enable"){ # format text according to action
			Out-ToHostAndFile "Enabled Successfully:`t`t`t$($Script:SuccessCount)"
		} else {
			Out-ToHostAndFile "Disabled Successfully:`t`t`t$($Script:SuccessCount)"
		}		
    }
    Out-ToHostAndFile "Already $($Action)d:`t`t`t$($Script:AlreadyEnabledOrDisabledCount)"
    Out-ToHostAndFile "Failed to $($Action):`t`t`t$($Script:FailedCount)"
    Out-ToHostAndFile "Not Compatible with HUB:`t`t$($Script:VMNotCompatibleCount)`n"
	if ($InputFilePath -ne "null"){
		$totalmachines = ($machines.count)
		Out-ToHostAndFile "VMs Inputted from file:`t`t`t$totalmachines"
	}
    Out-ToHostAndFile "Total VMs Processed:`t`t`t$($TotalVMsProcessed)"
    Out-ToHostAndFile "=====================================================================`n`n"
}


#######################################################
# Start PowerShell Script
#######################################################

# Define max threads for VM-level runspaces
$maxVMThreads = 7

<# $SubRunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $maxSubThreads)
$SubJobs = @()
$SubRunspacePool.Open()

$RGRunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $maxRGThreads)
$RGJobs = @()
$RGRunspacePool.Open() #>

# initialize logging paths
Set-OutputLogFiles

[string]$DateTimeNow = Get-Date -Format "dd/MM/yyyy - HH:mm:ss"
Out-ToHostAndFile "=====================================================================`n"
if ($Action -eq "enable"){ # format text according to action
	Out-ToHostAndFile "$($DateTimeNow) - Enable AHUB LicenseType Script Starting...`n"
} else {
	Out-ToHostAndFile "$($DateTimeNow) - Disable AHUB LicenseType Script Starting...`n"
}
Out-ToHostAndFile "====================================================================="
Out-ToHostAndFile " "

Get-AzureConnection
Run-AHUBToggleProcess 

[string]$DateTimeNow = get-date -Format "dd/MM/yyyy - HH:mm:ss"
Out-ToHostAndFile "=====================================================================`n"
if ($Action -eq "enable"){ # format text according to action
	Out-ToHostAndFile "$($DateTimeNow) - Enable AHUB LicenseType Script Complete`n"
} else {
	Out-ToHostAndFile "$($DateTimeNow) - Disable AHUB LicenseType Script Complete`n"
}
Out-ToHostAndFile "====================================================================="
Out-ToHostAndFile " "


