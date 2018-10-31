######################################
#                                    #
#        Custom Image Build     #
#         W10x64 - NoCopyProfile     #
#              David Dyer            #
#              07/25/2017            #
#                                    #
######################################


######################################

           #Copy-Items#

######################################

# - Copy background to wallpaper directory
New-Item C:\Windows\Web\Wallpaper\County -ItemType Directory -Force
Copy-Item -Path backgroundDefault.jpg -Destination C:\Windows\Web\Wallpaper\County -Force
Copy-Item -Path 'CountyUser.theme' -Destination C:\Windows\Resources\Themes -Force

# - Copy account picture to replace Windows default
Rename-Item 'C:\ProgramData\Microsoft\User Account Pictures\user.bmp' -NewName 'user.bak' -Force
Copy-Item -Path user.bmp -Destination 'C:\ProgramData\Microsoft\User Account Pictures' -Force

# - Event Log Cleaning
Copy-Item -Path 'ClearAllLogs.bat' -Destination 'C:\Windows' -Force

# - Documents
Copy-Item -Path 'Documents.lnk' -Destination 'C:\Users\Public\Desktop' -Force

# - Copy Internet Explorer link for Taskbar, Start Menu, and Desktop
Copy-Item -Path "Internet Explorer.lnk" -Destination "$env:ALLUSERSPROFILE\Microsoft\Windows\Start Menu\Programs\Accessories"
Copy-Item -Path "Internet Explorer.lnk" -Destination "$env:PUBLIC\Desktop"

######################################

          #Remove Items#

######################################

# - Remove Mixed Reality Portal so it never gets installed for new users
Remove-Item -Path "$env:SystemRoot\SystemApps\Microsoft.Windows.HolographicFirstRun_cw5n1h2txyewy" -Recurse -Force
Remove-Item -Path "$env:SystemRoot\SystemApps\Microsoft.PPIProjection_cw5n1h2txyewy" -Recurse -Force

######################################

          #Service Changes#

######################################

# - Change Remote Registry Service to Automatic Startup
Set-Service -Name RemoteRegistry -StartupType Automatic

# - Change Smart Card Service to Disabled
#Set-Service -Name SCardSvr -StartupType Disabled

# - Change Windows Remote Management Service to Automatic
Set-Service -Name WinRM -StartupType Automatic

######################################

          #Registry Changes#

######################################

# START **HKLM**

# - Disable SMB1
Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters" -Name SMB1 -Value 0 -Force

# - Disable Cortana
New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search" -Name AllowCortana -PropertyType DWORD -Value 0

# - Hide All Quick Actions
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Shell\ActionCenter\Quick Actions" -Name PinnedQuickActionSlotCount -Value 0 -Force

# - Hide Gaming and Mixed Reality Settings Pages
New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" -Name SettingsPageVisibility -PropertyType STRING -Value "hide:gaming-gamebar;gaming-gamedvr;gaming-broadcasting;gaming-gamemode;holographic;holographic-audio"

# - Disable Windows 10 Consumer Experience
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" -Name DisableWindowsConsumerFeatures -Value 1 -Force

# END **HKLM**

# START **HKCU**
# - Load Default User Hive
reg load 'HKLM\TempUser' C:\Users\Default\NTUSER.DAT

# - Create PSDrive at Default User Hive
New-PSDrive -Name HKDefaultUser -PSProvider Registry -Root HKLM\TempUser

# - Set Trusted and Intranet sites for Default User
# - Trusted Sites
New-Item -Path 'HKDefaultUser:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\colorado.gov'
New-Item -Path 'HKDefaultUser:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\colorado.gov\www'
New-ItemProperty -Path 'HKDefaultUser:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\colorado.gov\www' -PropertyType DWORD -Name https -Value 2
# - Intranet Sites
New-Item -Path 'HKDefaultUser:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\bouldercounty.org'
New-Item -Path 'HKDefaultUser:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\bouldercounty.org\www'
New-ItemProperty -Path 'HKDefaultUser:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\bouldercounty.org\www' -PropertyType DWORD -Name https -Value 1

# - Clear Browsing History on Exit
New-Item -Path 'HKDefaultUser:\Software\Microsoft\Internet Explorer' -Name Privacy -ItemType Directory
Set-ItemProperty -Path 'HKDefaultUser:\Software\Microsoft\Internet Explorer\Privacy' -Name ClearBrowsingHistoryOnExit -Value 1 -Force

# - Empty Temporary Internet Files Folder on Exit
Set-ItemProperty -Path 'HKDefaultUser:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Cache' -Name Persistent -Value 0 -Force

# - Disable People in Taskbar
Set-ItemProperty -Path 'HKDefaultUser:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name PeopleBand -Value 0 -Force

# - Enable Desktop "This PC" Icon
New-Item -Path 'HKDefaultUser:\Software\Microsoft\Windows\CurrentVersion\Explorer' -Name HideDesktopIcons -ItemType Directory
New-Item -Path 'HKDefaultUser:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons' -Name NewStartPanel -ItemType Directory
New-ItemProperty -Path 'HKDefaultUser:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel' -PropertyType DWORD -Name '{20D04FE0-3AEA-1069-A2D8-08002B30309D}' -Value 0

# - Disable "Occasionally show suggestions in Start"
New-ItemProperty -Path 'HKDefaultUser:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager' -Name 'SubscribedContent-338388Enabled' -PropertyType DWORD -Value 0 -Force

# - Disable "Get tips, tricks, and suggestions as you use Windows"
New-ItemProperty -Path 'HKDefaultUser:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager' -Name 'SubscribedContent-310093Enabled' -PropertyType DWORD -Value 0 -Force

# - Disable "Show me the Windows welcome experience"
New-ItemProperty -Path 'HKDefaultUser:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager' -Name 'SubscribedContent-338389Enabled' -PropertyType DWORD -Value 0 -Force

# - Disable Live Tile Notifications
New-Item -Path 'HKDefaultUser:\Software\Policies\Microsoft\Windows\CurrentVersion' -Name 'PushNotifications' -ItemType Directory -Force
New-ItemProperty -Path 'HKDefaultUser:\Software\Policies\Microsoft\Windows\CurrentVersion\PushNotifications' -Name 'NoTileApplicationNotification' -PropertyType DWORD -Value 0 -Force

# - Remove Default User Drive
Remove-PSDrive HKDefaultUser

# - Clean handles connected to user hive
[gc]::Collect()

# - Unload Registry Hive
reg unload 'HKLM\TempUser'

# - Clean handles connected to user hive
[gc]::Collect()

# END **HKCU**

#########################################
          
  #Enable Remote Desktop and PSRemoting #

#########################################

(Get-WmiObject Win32_TerminalServiceSetting -Namespace root\cimv2\TerminalServices).SetAllowTsConnections(1,1) | Out-Null
(Get-WmiObject -Class "Win32_TSGeneralSetting" -Namespace root\cimv2\TerminalServices -Filter "TerminalName='RDP-tcp'").SetUserAuthenticationRequired(1) | Out-Null

#################################

  #Customize Windows Start Menu

#################################

Import-StartLayout -LayoutPath 'LayoutModification.xml' -MountPath "$env:SystemDrive\"