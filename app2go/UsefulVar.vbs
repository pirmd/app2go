' -----------------------------------------------------------------------------
'         NAME: app2go_UsefulEnv.vbs
'  DESCRIPTION: determine and make available useful vars
'      CREATED: 2012.08.08 / REVISION: 2014.07.30 - 16:06
' -----------------------------------------------------------------------------

LauncherDrive  = FSO.GetDriveName(WScript.ScriptFullName)
LauncherFolder = FSO.GetParentFolderName(WScript.ScriptFullName)
LauncherParentFolder = FSO.GetParentFolderName(LauncherFolder)

MyDocsOnHost = WshShell.SpecialFolders("MyDocuments")
DesktopOnHost = WshShell.SpecialFolders("Desktop")
SendToOnHost = WshShell.SpecialFolders("SendTo")

AppdirOnHost = WshShell.ExpandEnvironmentStrings("%APPDATA%")
TempOnHost = WshShell.ExpandEnvironmentStrings("%TEMP%")
HostName = WshShell.ExpandEnvironmentStrings("%COMPUTERNAME%")

HKCU="HKEY_CURRENT_USER\"
HKCU_CLASSES=HKCU&"SOFTWARE\CLASSES\"
