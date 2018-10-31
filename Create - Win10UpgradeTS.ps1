<#
.SYNOPSIS
    This script creates an in-place upgrade Task Sequence in Microsoft System Center to be used as a template.
.DESCRIPTION
    Script needs to be ran with an account that has appropriate SCCM permissions.
    The script contains some lines to be modified to fit the environment which it is running,
    including Task Sequence Name, Operating System Upgrade Object, Driver Package Name, Driver Package Model Query.
    Modify the lines between "Start Modify Area" and "End Modify Area" accordingly.
.NOTES
    Created on: 7/10/2018
    Created by: Brian Embree
    Last modified: 8/23/2018
    Last modified by: Brian Embree
#>

#Load Configuration Manager PowerShell Module
Import-module ($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5) + '\ConfigurationManager.psd1')

#Get SiteCode
$SiteCode = Get-PSDrive -PSProvider CMSITE
Set-location $SiteCode":"

#_____Start Modify Area_____#

# Set Task Sequence Name
$TaskSequenceName = 'Windows10Upgrade'

# Get Operating System Upgrade object (must match Operating System Upgrade name in SCCM)
$OperatingSystemUpgradePackage = Get-CMOperatingSystemUpgradePackage -Name "Windows 10"

# Driver package name for initial test (must match driver package name in SCCM)
$DriverPackage1Name = 'Dell - E8 Latitude Family - Win10x64'

# Computer Model Query for Model that fits the selected Driver Package, change out 'HP EliteBook 840G3' with appropriate Model (wildcard capable)
$DriverPackage1Query = New-CMTaskSequenceStepConditionQueryWMI -Namespace 'root\cimv2' -Query "SELECT * FROM Win32_ComputerSystem WHERE Model like 'HP EliteBook 840G3'"

#_____End Modify Area_____#

# Create new Task Sequence using a Splatting Hash-Table, using the generic enterprise KMS key
$TaskSequenceArgs = @{
    UpgradeOperatingSystem = $true
    Name = $TaskSequenceName
    UpgradePackageId = $OperatingSystemUpgradePackage.PackageID
    ProductKey = 'NPPR9-FWDCX-D2C8J-H872K-2YT43'
}
New-CMTaskSequence @TaskSequenceArgs | Out-Null

# Set Task Sequence Variable to active Task Sequence
$TaskSequence = Get-CMTaskSequence -Name $TaskSequenceName

# Define built in suggested groups provided by Microsoft under "Prepare for Upgrade"
$PrepareForUpgrade = @()
$PrepareForUpgrade += $BatteryChecks = New-CMTaskSequenceGroup -Name 'Battery Checks' -Description 'Add steps in this group to check whether the computer is using battery, or wired power'
$PrepareForUpgrade += $NetworkConnectionChecks = New-CMTaskSequenceGroup -Name 'Network/Wired Connection Checks' -Description 'Add steps in this group to check whether the computer is connected to a network, and is not using a wireless connection'
$PrepareForUpgrade += $RemoveIncompatibleApps = New-CMTaskSequenceGroup -Name 'Remove Incompatible Applications' -Description 'Add steps in this group to remove any applications that are incompatible with this version of Windows 10'
$PrepareForUpgrade += $RemoveIncompatibleDrivers = New-CMTaskSequenceGroup -Name 'Remove Incompatible Drivers' -Description 'Add steps in this group to remove any drivers that are incompatible with this version of Windows 10'
$PrepareForUpgrade += $RemoveThirdPartysecurity = New-CMTaskSequenceGroup -Name 'Remove/Suspend Third-party Security' -Description 'Add steps in this group to remove or suspend third-party security programs, such as antivirus'

# Define built in suggested groups provided by Microsoft under "Post-Processing"
$PostProcessing = @()
$PostProcessing += $ApplySetupBasedDrivers = New-CMTaskSequenceGroup -Name 'Apply setup-based drivers' -Description 'Add steps in this group to install setup-based drivers (.exe) from packages'
$PostProcessing += $InstallThirdPartySecurity = New-CMTaskSequenceGroup -Name 'Install/Enable Third-party Security' -Description 'Add steps in this group to install or enable third-party security programs, such as antivirus'
$PostProcessing += $SetWindowsDefaultApps = New-CMTaskSequenceGroup -Name 'Set Windows Default Apps and Associations' -Description 'Add steps in this group to set Windows default apps and file associations'
$PostProcessing += $ApplyCustomizationsAndPersonalizations = New-CMTaskSequenceGroup -Name 'Apply Customizations and Personalizations' -Description 'Add steps in this group to apply Start menu customizations, such as organizing program groups'

# Define built in suggested groups provided by Microsoft under "Run Actions on Failure"
$FailureActions = @()
$FailureActions += $CollectLogs = New-CMTaskSequenceGroup -Name 'Collect Logs' -Description 'Save Windows Setup and Task Sequence logs to a Network Share'
$FailureActions += $RunDiagnosticTools = New-CMTaskSequenceGroup -Name 'Run Diagnostic Tools' -Description 'Run Diagnostic Tools'

# Create built in group Run Actions on Failure
$RunActionsFailureCondition = New-CMTaskSequenceStepConditionVariable -OperatorType NotEquals -ConditionVariableName _SMSTSOSUpgradeActionReturnCode -ConditionVariableValue 0
$RunActionsonFailureGroup = New-CMTaskSequenceGroup -Name 'Run Actions on Failure' -Description 'Run tasks if OS upgrade fails to gather logs and run diagnostic tools' -Condition $RunActionsFailureCondition
Add-CMTaskSequenceStep -TaskSequenceId $TaskSequence.PackageID -Step $RunActionsonFailureGroup -InsertStepStartIndex 4

# Create built in suggested groups provided by Microsoft under "Prepare for Upgrade"
foreach($Prep in $PrepareForUpgrade){
Set-CMTaskSequenceGroup -TaskSequenceId $TaskSequence.PackageID -AddStep $Prep -StepName 'Prepare for Upgrade'
}

# Create built in suggested groups provided by Microsoft under "Post-Processing"
foreach($Post in $PostProcessing){
Set-CMTaskSequenceGroup -TaskSequenceId $TaskSequence.PackageID -AddStep $Post -StepName 'Post-Processing'
}

# Create built in suggested groups provided by Microsoft under "Run Actions on Failure"
foreach($Action in $FailureActions){
Set-CMTaskSequenceGroup -TaskSequenceId $TaskSequence.PackageID -AddStep $Action -StepName 'Run Actions on Failure'
}

# Create and add Upgrade Assessment step for testing the upgrade prior to actual upgrade
$UpgradeAssessmentStep = New-CMTaskSequenceStepUpgradeOperatingSystem -UpgradePackage $OperatingSystemUpgradePackage -ProductKey 'NPPR9-FWDCX-D2C8J-H872K-2YT43' -Name 'Upgrade Assessment' -ScanOnly $true -IgnoreMessage $true
Set-CMTaskSequenceGroup -TaskSequenceId $TaskSequence.PackageID -StepName 'Prepare for Upgrade' -AddStep $UpgradeAssessmentStep -InsertStepStartIndex 1

# Modify Upgrade Operating System step to Ignore Dismissable Messages, add step and condition
Set-CMTaskSequenceStepUpgradeOperatingSystem -TaskSequenceId $TaskSequence.PackageID -StepName 'Upgrade the Operating System' -IgnoreMessage $true
$UpgradeTheOperatingSystemCondition = New-CMTaskSequenceStepConditionVariable -OperatorType Equals -ConditionVariableName '_SMSTSOSUpgradeActionReturnCode' -ConditionVariableValue '3247440400'
Set-CMTaskSequenceGroup -TaskSequenceId $TaskSequence.PackageID -StepName 'Upgrade the Operating System' -AddCondition $UpgradeTheOperatingSystemCondition

# Create and add Download Drivers group
$DownloadDriversGroup = New-CMTaskSequenceGroup -Name 'Download Drivers'
Set-CMTaskSequenceGroup -TaskSequenceId $TaskSequence.PackageID -AddStep $DownloadDriversGroup -StepName 'Upgrade the Operating System' -InsertStepStartIndex 0

# Get Driver Package, create and add Download Driver Package Content step, shorten name if too long for step name
$DriverPackage1 = Get-CMDriverPackage -Name "$DriverPackage1Name"
if($DriverPackage1Name.Length -gt 23){
    $DriverPackage1NameShort = $DriverPackage1Name.Substring(0,23); $DriverPackage1NameShort = $DriverPackage1NameShort + '...'
}
$DownloadDriverPackage1Content = New-CMTaskSequenceStepDownloadPackageContent -AddPackage $DriverPackage1 -Name "Download $($DriverPackage1NameShort) Driver Package" -LocationOption CustomPath -Path '%_SMSTSDataPath%\Win10Drivers' -Condition $DriverPackage1Query
Set-CMTaskSequenceGroup -TaskSequenceId $TaskSequence.PackageID -StepName 'Download Drivers' -AddStep $DownloadDriverPackage1Content

# Create and add Use Drivers if Available step and condition
$UseDriversIfAvailableCondition = New-CMTaskSequenceStepConditionFolder -FolderPath '%_SMSTSMDataPath%\Win10Drivers'
$UseDriversIfAvailableStep = New-CMTaskSequenceStepSetVariable -TaskSequenceVariable 'OSDUpgradeStagedContent' -TaskSequenceVariableValue '%_SMSTSMDataPath%\Win10Drivers' -Name 'Install Drivers only if Available' -Condition $UseDriversIfAvailableCondition
Set-CMTaskSequenceGroup -TaskSequenceId $TaskSequence.PackageID -StepName 'Download Drivers' -AddStep $UseDriversIfAvailableStep

# Create High Performance Step
$SetPowerSchemeHighPerfStep = New-CMTaskSequenceStepRunCommandLine -CommandLine 'PowerCfg.exe /s 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c' -Name 'Set PowerScheme to High Performance'
Set-CMTaskSequenceGroup -TaskSequenceId $TaskSequence.PackageID -AddStep $SetPowerSchemeHighPerfStep -StepName 'Prepare for Upgrade' -InsertStepStartIndex 0

# Create and add Disable BitLocker group and condition
$DisableBitLockerCondition = New-CMTaskSequenceStepConditionQueryWMI -Namespace 'root\cimv2\Security\MicrosoftVolumeEncryption' -Query "SELECT * FROM Win32_EncryptableVolume WHERE DriveLetter = 'C:' and ProtectionStatus = '1'"
$DisableBitLockerGroup = New-CMTaskSequenceGroup -Name 'Disable BitLocker' -Condition $DisableBitLockerCondition -Disable
Set-CMTaskSequenceGroup -TaskSequenceId $TaskSequence.PackageID -AddStep $DisableBitLockerGroup -StepName 'Prepare For Upgrade' -InsertStepStartIndex 3

# Create and add Set BitLocker Status step
$SetOSDBitLockerStatusStep = New-CMTaskSequenceStepSetVariable -TaskSequenceVariable 'OSDBitLockerStatus' -TaskSequenceVariableValue 'Protected' -Name 'Set OSD BitLocker Status'
Set-CMTaskSequenceGroup -TaskSequenceId $TaskSequence.PackageID -StepName 'Disable BitLocker' -AddStep $SetOSDBitLockerStatusStep

# Create and add Disable BitLocker step
$DisableBitlockerStep = New-CMTaskSequenceStepRunCommandLine -CommandLine 'manage-bde -protectors -disable C: -RC 0' -Name 'Disable BitLocker (Multiple Reboots)'
Set-CMTaskSequenceGroup -TaskSequenceId $TaskSequence.PackageID -StepName 'Disable BitLocker' -AddStep $DisableBitlockerStep

# Create Update BIOS/Firmware Group
$UpdateBIOSFirmwareGroup = New-CMTaskSequenceGroup -Name 'Update BIOS/Firmware'
Set-CMTaskSequenceGroup -TaskSequenceId $TaskSequence.PackageID -AddStep $UpdateBIOSFirmwareGroup -StepName 'Prepare for Upgrade' -InsertStepStartIndex 9

# Create and add Re-Enable BitLocker group
$ReEnableBitLockerGroup = New-CMTaskSequenceGroup -Name 'Re-Enable BitLocker' -Disable
Add-CMTaskSequenceStep -TaskSequenceId $TaskSequence.PackageID -Step $ReEnableBitLockerGroup -InsertStepStartIndex 3

# Create and add Enable BitLocker step and condition
$EnableBitLockerCondition = New-CMTaskSequenceStepConditionVariable -ConditionVariableName 'OSDBitLockerStatus' -OperatorType Equals -ConditionVariableValue 'protected'
$EnableBitLockerStep = New-CMTaskSequenceStepEnableBitLocker -Name 'Enable BitLocker' -TpmOnly -Condition $EnableBitLockerCondition
Set-CMTaskSequenceGroup -TaskSequenceId $TaskSequence.PackageID -StepName 'Re-Enable BitLocker' -AddStep $EnableBitLockerStep

# Create Set TS Resiliency Group
$SetTSResiliencyGroup = New-CMTaskSequenceGroup -Name 'Set TS Resiliency'
Add-CMTaskSequenceStep -TaskSequenceId $TaskSequence.PackageID -Step $SetTSResiliencyGroup -InsertStepStartIndex 0

# Create SMSTS Download Retry Count Variable Step
$SetSMSTSDownloadRetryCountStep = New-CMTaskSequenceStepSetVariable -TaskSequenceVariable 'SMSTSDownloadRetryCount' -TaskSequenceVariableValue '5' -Name 'Set SMSTSDownloadRetryCount Variable'
Set-CMTaskSequenceGroup -TaskSequenceId $TaskSequence.PackageID -AddStep $SetSMSTSDownloadRetryCountStep -StepName 'Set TS Resiliency'

# Create SMSTS Download Retry Delay Variable Step
$SetSMSTSDownloadRetryDelayStep = New-CMTaskSequenceStepSetVariable -TaskSequenceVariable 'SMSTSDownloadRetryDelay' -TaskSequenceVariableValue '30' -Name 'Set SMSTSDownloadRetryDelay Variable'
Set-CMTaskSequenceGroup -TaskSequenceId $TaskSequence.PackageID -AddStep $SetSMSTSDownloadRetryDelayStep -StepName 'Set TS Resiliency'

# Create SMSTS MP List Request Timeout Variable Step
$SetSMSTSMPListRequestTimeoutStep = New-CMTaskSequenceStepSetVariable -TaskSequenceVariable 'SMSTSMPListRequestTimeout' -TaskSequenceVariableValue '300000' -Name 'Set SMSTSMPListRequestTimeout Variable'
Set-CMTaskSequenceGroup -TaskSequenceId $TaskSequence.PackageID -AddStep $SetSMSTSMPListRequestTimeoutStep -StepName 'Set TS Resiliency'

# Create Balanced Performance Step
$SetPowerSchemeBalancedStep = New-CMTaskSequenceStepRunCommandLine -CommandLine 'PowerCfg.exe /s 381b4222-f694-41f0-9685-ff5bb260df2e' -Name 'Set PowerScheme to Balanced'
Set-CMTaskSequenceGroup -TaskSequenceId $TaskSequence.PackageID -AddStep $SetPowerSchemeBalancedStep -StepName 'Apply Customizations and Personalizations'

# Create Remove Office Applications Group
$RemoveOfficeAppsGroup = New-CMTaskSequenceGroup -Name 'Remove Office Applications (OffScrub)'
Set-CMTaskSequenceGroup -TaskSequenceId $TaskSequence.PackageID -AddStep $RemoveOfficeAppsGroup -StepName 'Post-Processing' -InsertStepStartIndex 1

# Create Install New Software Group
$InstallNewSoftwareGroup = New-CMTaskSequenceGroup -Name 'Install New Software'
Set-CMTaskSequenceGroup -TaskSequenceId $TaskSequence.PackageID -AddStep $InstallNewSoftwareGroup -StepName 'Post-Processing' -InsertStepStartIndex 3

# Create Configure UEFI Group
$ConfigureUEFICondition = New-CMTaskSequenceStepConditionVariable -ConditionVariableName '_SMSTSBootUEFI' -OperatorType NotEquals -ConditionVariableValue 'True'
$ConfigureUEFIGroup = New-CMTaskSequenceGroup -Name 'Configure UEFI' -Condition $ConfigureUEFICondition
Set-CMTaskSequenceGroup -TaskSequenceId $TaskSequence.PackageID -AddStep $ConfigureUEFIGroup -StepName 'Post-Processing' -InsertStepStartIndex 6

# Create Remove Default Apps Win 10 Placeholder Step
$RemoveDefaultAppsWin10Step = New-CMTaskSequenceStepRunCommandLine -Name 'Remove Default Apps Win 10' -CommandLine 'Placeholder for RUN POWERSHELL SCRIPT STEP to run RemoveDefaultAppsWin10.ps1' -Description 'Placeholder for RUN POWERSHELL SCRIPT STEP to run RemoveDefaultAppsWin10.ps1' -Disable
Set-CMTaskSequenceGroup -TaskSequenceId $TaskSequence.PackageID -AddStep $RemoveDefaultAppsWin10Step -StepName 'Set Windows Default Apps and Associations'

# Create Set Default Applications Placeholder Step
$SetDefaultApplicationsStep = New-CMTaskSequenceStepRunCommandLine -Name 'Set Default Applications' -CommandLine 'Dism.exe /Online /Import-DefaultAppAssociations:DefaultAppAssoc.xml' -Description 'Add DefaultAppAssoc.xml file package content' -Disable
Set-CMTaskSequenceGroup -TaskSequenceId $TaskSequence.PackageID -AddStep $SetDefaultApplicationsStep -StepName 'Set Windows Default Apps and Associations'

# Create Modify Start Menu and Taskbar Step
$ModifyStartMenuandTaskbarStep = New-CMTaskSequenceStepRunCommandLine -Name 'Modify Start Menu and Taskbar' -CommandLine 'Placeholder for RUN POWERSHELL SCRIPT STEP to run ImportStartMenu.ps1' -Description 'Placeholder for RUN POWERSHELL SCRIPT STEP to run ImportStartMenu.ps1' -Disable
Set-CMTaskSequenceGroup -TaskSequenceId $TaskSequence.PackageID -AddStep $ModifyStartMenuandTaskbarStep -StepName 'Apply Customizations and Personalizations' -InsertStepStartIndex 0

# Create Apply Local Group Policy Settings Step
$ApplyLocalGroupPolicySettingsStep = New-CMTaskSequenceStepRunCommandLine -Name 'Apply Local Group Policy Settings' -CommandLine 'LGPO.exe /g .\' -Description 'Add LGPO.exe and settings file package content' -Disable
Set-CMTaskSequenceGroup -TaskSequenceId $TaskSequence.PackageID -AddStep $ApplyLocalGroupPolicySettingsStep -StepName 'Apply Customizations and Personalizations' -InsertStepStartIndex 1

# Create First Restart Computer Step for Configure UEFI Group
$FirstRestartComputerStep = New-CMTaskSequenceStepReboot -Name 'Restart Computer' -RunAfterRestart WinPE -NotificationMessage 'A new Microsoft Windows Operating system is being installed. The computer must restart to continue.' -MessageTimeout 15 -Disable
Set-CMTaskSequenceGroup -TaskSequenceId $TaskSequence.PackageID -AddStep $FirstRestartComputerStep -StepName 'Configure UEFI'

# Create Convert MBR to GPT Step for Configure UEFI Group
$ConvertMBRtoGPTStep = New-CMTaskSequenceStepRunCommandLine -Name 'Convert MBR to GPT' -CommandLine 'mbr2gpt.exe /convert' -Disable
Set-CMTaskSequenceGroup -TaskSequenceId $TaskSequence.PackageID -AddStep $ConvertMBRtoGPTStep -StepName 'Configure UEFI'

# Create Convert BIOS to UEFI - HP Step
$ConvertBIOStoUEFICondition = New-CMTaskSequenceStepConditionQueryWMI -Namespace root\cimv2 -Query 'SELECT * FROM Win32_ComputerSystem WHERE Manufacturer LIKE "%HP%"'
$ConvertBIOStoUEFIStep = New-CMTaskSequenceStepRunCommandLine -Name 'Convert BIOS to UEFI - HP' -CommandLine 'BiosConfigUtility64.exe /setconfig:UEFI.txt /cspwdfile:password.bin' -Description 'Add Vendor Bios Config Utility package content and modify step accordingly' -Condition $ConvertBIOStoUEFICondition -Disable
Set-CMTaskSequenceGroup -TaskSequenceId $TaskSequence.PackageID -AddStep $ConvertBIOStoUEFIStep -StepName 'Configure UEFI'

# Create Last Restart Computer Step for Configure UEFI Group
$LastRestartComputerStep = New-CMTaskSequenceStepReboot -Name 'Restart Computer' -RunAfterRestart HardDisk -NotificationMessage '' -Disable
Set-CMTaskSequenceGroup -TaskSequenceId $TaskSequence.PackageID -AddStep $LastRestartComputerStep -StepName 'Configure UEFI'
