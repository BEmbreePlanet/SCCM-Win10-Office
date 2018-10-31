<#
	Created on: 5/4/18
	Created by: Brian Embree
	Last modified: 5/25/2018
	Last modified by: Brian Embree
	Description: Creates initial device collections and folders for new SCCM environment.
#>

#Load Configuration Manager PowerShell Module
Import-module ($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5) + '\ConfigurationManager.psd1')

#Get SiteCode
$SiteCode = Get-PSDrive -PSProvider CMSITE
Set-location $SiteCode":"

#Error Handling and output
Clear-Host
$ErrorActionPreference= 'SilentlyContinue'
$Error1 = 0

#Create Collection Refresh Schedule
$Date = Get-Date
$Schedule = New-CMSchedule -Start "$Date" -RecurInterval Days -RecurCount 7

$SoftwareUpdateFolder = "Software Update Management"

#Create Device Collection Folders
New-Item -Name $SoftwareUpdateFolder -Path "$($SiteCode):\DeviceCollection"

#####Collection Name and Query to be created#####
$Collection1 = @{Name = 'All Workstations'; Query = 'select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System where SMS_R_System.OperatingSystemNameandVersion like "%Workstation%"'}
$Collection2 = @{Name = 'All Servers Non Domain Controllers'; Query = 'select SMS_R_System.ResourceId, SMS_R_System.ResourceType, SMS_R_System.Name, SMS_R_System.SMSUniqueIdentifier, SMS_R_System.ResourceDomainORWorkgroup, SMS_R_System.Client from  SMS_R_System where SMS_R_System.OperatingSystemNameandVersion like "%Server%" and SMS_R_System.PrimaryGroupID != 516'}
$Collection3 = @{Name = 'All Domain Controllers'; Query = 'select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System     WHERE SMS_R_System.PrimaryGroupID = "516"'}
$Collection4 = @{Name = 'SUM - Excluded Workstations'; Query = ''}
$Collection5 = @{Name = 'SUM - Pilot 1'}
$Collection6 = @{Name = 'SUM - Pilot 2'}
$Collection7 = @{Name = 'SUM - Production - Workstations'}

#####Any Collection that needs to be limited to All Systems#####
$LimitingCollection1 = 'All Systems'
$LimitToLimitingCollection1 = @($Collection1, $Collection2, $Collection3)

#####Any Software Update Management Collections, no query, limited to Collection1 "All Workstations" and moved to $SoftwareUpdateFolder.##### 
#####Collection Include and exclude added after loop#####
$LimitingCollection2 = $Collection1.Name
$LimitToLimitingCollection2 = @($Collection4, $Collection5, $Collection6, $Collection7)


#Loop to create collections based on limiting collection
try{
    foreach($Collection in $LimitToLimitingCollection1){
        New-CMDeviceCollection -Name $Collection.Name -LimitingCollectionName $LimitingCollection1 -RefreshType Both -RefreshSchedule $Schedule | Out-Null
        Add-CMDeviceCollectionQueryMembershipRule -CollectionName $Collection.Name -QueryExpression $Collection.Query -RuleName $Collection.Name
        Write-Host -ForegroundColor Green "*****$($Collection.Name) Created*****"
    }
    foreach($Collection in $LimitToLimitingCollection2){
        New-CMDeviceCollection -Name $Collection.Name -LimitingCollectionName $LimitingCollection2 -RefreshType Both -RefreshSchedule $Schedule | Out-Null
        Write-Host -ForegroundColor Green "*****$($Collection.Name) Created*****"
        Get-CMDeviceCollection -Name $Collection.Name | Move-CMObject -FolderPath "$($SiteCode):\DeviceCollection\Software Update Management"
        Write-Host -ForegroundColor Yellow "*****$($Collection.Name) Moved to $($SoftwareUpdateFolder) "
    }

Add-CMDeviceCollectionIncludeMembershipRule -CollectionName $Collection7.Name -IncludeCollectionName $Collection1.Name
Add-CMDeviceCollectionExcludeMembershipRule -CollectionName $Collection7.Name -ExcludeCollectionName $Collection4.Name
Add-CMDeviceCollectionExcludeMembershipRule -CollectionName $Collection7.Name -ExcludeCollectionName $Collection5.Name
Add-CMDeviceCollectionExcludeMembershipRule -CollectionName $Collection7.Name -ExcludeCollectionName $Collection6.Name
}
Catch{$Error1 = 1}

Finally{
    If ($Error1 -eq 1){
        Write-host "-----------------"
        Write-host -ForegroundColor Red "Script has already been run or a collection name already exist."
        Write-host "-----------------"
        Pause
    }
    Else{
        Write-host "-----------------"
        Write-Host -ForegroundColor Green "Script execution completed without errors. Collections created sucessfully"
        Write-host "-----------------"
        Pause
        }
        }

