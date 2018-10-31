# Note #1: 
# To service a newer version of WinPE than the OS you are servicing from, for example service Windows 10 v1709 
# from a Windows Server 2016 server, you need a newer DISM version.
# Solution, simply install the latest Windows ADK 10, and use DISM from that version
#
# Note #2:
# If your Windows OS already have a newer version of dism, uncomment the below line, and comment out line 11 and 12
$DISMFile = 'dism.exe'

# Configuring the script to use the Windows ADK 10 version of DISM
# $DISMFile = 'C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\DISM\dism.exe'
# If (!(Test-Path $DISMFile)){ Write-Warning "DISM in Windows ADK not found, aborting..."; Break }

# Set additional parameters
$ISO = "C:\Users\Bembree\OneDrive - Planet Technologies, Inc\Lab Files\en_windows_10_business_editions_version_1803_updated_march_2018_x64_dvd_12063333.iso"
$ServicingUpdate = "C:\Users\Bembree\OneDrive - Planet Technologies, Inc\Lab Files\7-31-2018 1803 Updates\windows10.0-kb4343669-x64_2a58320e44d3ff803bc7016b5d02f3e85482b46f.msu"
$AdobeFlashUpdate = "C:\Users\Bembree\OneDrive - Planet Technologies, Inc\Lab Files\7-31-2018 1803 Updates\windows10.0-kb4338832-x64_961806f0d9d52c7df32346349c5101f1ab2032a0.msu"
$MonthlyCU = "C:\Users\Bembree\OneDrive - Planet Technologies, Inc\Lab Files\7-31-2018 1803 Updates\windows10.0-kb4340917-x64_ed5f586fa9d7064d61e45c523f20ba7fc1ff142a.msu"
$ImageMountFolder = "C:\Mount_Image"
$BootImageMountFolder = "C:\Mount_BootImage"
$WIMImageFolder = "C:\WIMs"
$TmpImage = "$WIMImageFolder\tmp_install.wim"
$TmpWinREImage = "$WIMImageFolder\tmp_winre.wim"
$RefImage = "$WIMImageFolder\install.wim"
$BootImage = "$WIMImageFolder\boot.wim"

New-Item -Path 'C:\' -Name Mount_Image -ItemType Directory
New-Item -Path 'C:\' -Name Mount_BootImage -ItemType Directory
New-Item -Path 'C:\' -Name WIMs -ItemType Directory

# Verify that files and folder exist
if (!(Test-Path -path $ISO)) {Write-Warning "Could not find Windows 10 ISO file. Aborting...";Break}
if (!(Test-Path -path $ServicingUpdate)) {Write-Warning "Could not find Servicing Update for Windows 10. Aborting...";Break}
if (!(Test-Path -path $AdobeFLashUpdate)) {Write-Warning "Could not find Adobe Flash Update for Windows 10. Aborting...";Break}
if (!(Test-Path -path $MonthlyCU)) {Write-Warning "Could not find Monthly Update for Windows 10. Aborting...";Break}
if (!(Test-Path -path $ImageMountFolder)) {New-Item -path $ImageMountFolder -ItemType Directory}
if (!(Test-Path -path $BootImageMountFolder)) {New-Item -path $BootImageMountFolder -ItemType Directory}
if (!(Test-Path -path $WIMImageFolder)) {New-Item -path $WIMImageFolder -ItemType Directory}

# Check Windows Version
$OSCaption = (Get-WmiObject win32_operatingsystem).caption
If ($OSCaption -like "Microsoft Windows 10*" -or $OSCaption -like "Microsoft Windows Server 2016*")
{
    # All OK
}
Else
{
    Write-Warning "$Env:Computername Oupps, you really should use Windows 10 or Windows Server 2016 when servicing Windows 10 offline"
    Write-Warning "$Env:Computername Aborting script..."
    Break
}

# Mount the Windows 10 ISO
Mount-DiskImage -ImagePath $ISO
$ISOImage = Get-DiskImage -ImagePath $ISO | Get-Volume
$ISODrive = [string]$ISOImage.DriveLetter+":"

# Export the Windows 10 Enterprise index to a new (temporary) WIM
Export-WindowsImage -SourceImagePath "$ISODrive\Sources\install.wim" -SourceName "Windows 10 Enterprise" -DestinationImagePath $TmpImage

# Mount the Windows 10 Enterprise image/index with the Optimize option (reduces initial mount time)
Mount-WindowsImage -ImagePath $TmpImage -Index 1 -Path $ImageMountFolder -Optimize

# Add the Updates to the Windows 10 Enterprise image
& $DISMFile /Image:$ImageMountFolder /Add-Package /PackagePath:$ServicingUpdate
& $DISMFile /Image:$ImageMountFolder /Add-Package /PackagePath:$AdobeFlashUpdate
& $DISMFile /Image:$ImageMountFolder /Add-Package /PackagePath:$MonthlyCU

# Cleanup the image BEFORE installing .NET to prevent errors
# Using the /ResetBase switch with the /StartComponentCleanup parameter of DISM.exe on a running version of Windows 10 removes all superseded versions of every component in the component store.
# https://docs.microsoft.com/en-us/windows-hardware/manufacture/desktop/clean-up-the-winsxs-folder#span-iddismexespanspan-iddismexespandismexe
& $DISMFile /Image:$ImageMountFolder /Cleanup-Image /StartComponentCleanup /ResetBase

# Add .NET Framework 3.5.1 to the Windows 10 Enterprise image 
& $DISMFile /Image:$ImageMountFolder /Enable-Feature /FeatureName:NetFx3 /All /LimitAccess /Source:"$ISODrive\sources\sxs"

# Move WinRE Image to temp location
Move-Item -Path $ImageMountFolder\Windows\System32\Recovery\winre.wim -Destination $TmpWinREImage

# Mount the temp WinRE Image
Mount-WindowsImage -ImagePath $TmpWinREImage -Index 1 -Path $BootImageMountFolder

# Add the Updates to the WinRE image 
& $DISMFile /Image:$BootImageMountFolder /Add-Package /PackagePath:$ServicingUpdate
& $DISMFile /Image:$BootImageMountFolder /Add-Package /PackagePath:$MonthlyCU

# Cleanup the WinRE image
& $DISMFile /Image:$BootImageMountFolder /Cleanup-Image /StartComponentCleanup /ResetBase 
    
# Dismount the WinRE image
DisMount-WindowsImage -Path $BootImageMountFolder -Save

# Export new WinRE wim back to original location
Export-WindowsImage -SourceImagePath $TmpWinREImage -SourceName "Microsoft Windows Recovery Environment (x64)" -DestinationImagePath $ImageMountFolder\Windows\System32\Recovery\winre.wim

# Dismount the Windows 10 Enterprise image
DisMount-WindowsImage -Path $ImageMountFolder -Save

# Export the Windows 10 Enterprise index to a new WIM (the export operation reduces the WIM size with about 400 - 500 MB)
Export-WindowsImage -SourceImagePath $TmpImage -SourceName "Windows 10 Enterprise" -DestinationImagePath $RefImage

# Remove the temporary WIM
if (Test-Path -path $TmpImage) {Remove-Item -Path $TmpImage -Force}
if (Test-Path -path $TmpWinREImage) {Remove-Item -Path $TmpWinREImage -Force}

# Mount index 2 of the Windows 10 boot image (boot.wim)
Copy-Item "$ISODrive\Sources\boot.wim" $WIMImageFolder
Attrib -r $BootImage 
Mount-WindowsImage -ImagePath $BootImage -Index 2 -Path $BootImageMountFolder

# Add the Updates to the boot image
& $DISMFile /Image:$BootImageMountFolder /Add-Package /PackagePath:$ServicingUpdate
& $DISMFile /Image:$BootImageMountFolder /Add-Package /PackagePath:$MonthlyCU

# Dismount the boot image
DisMount-WindowsImage -Path $BootImageMountFolder -Save

# Dismount the Windows 10 ISO
Dismount-DiskImage -ImagePath $ISO 

#Remove Mount Folders
Remove-Item $ImageMountFolder
Remove-Item $BootImageMountFolder
