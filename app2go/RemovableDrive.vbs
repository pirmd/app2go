' -----------------------------------------------------------------------------
'         NAME: RemovableDrive.vbs
'  DESCRIPTION: helpers to access removable drive
'      CREATED: 2014.03.07 / REVISION: 2014.10.30 - 13:42
' -----------------------------------------------------------------------------

' -----------------------------------------------------------------------------
'        NAME: GetDriveByVolumeName
' DESCRIPTION: Get the drive name (H:, P:,...) of a disk given its volum name
'     PARAM 1: VolumeName to search for
'      OUTPUT: Drive Name or empty string
' -----------------------------------------------------------------------------
Public Function GetDriveByVolumeName(VolumeName)
    WMIQuery = "SELECT * from Win32_LogicalDisk WHERE VolumeName =""" & VolumeName & """"
    Set Drives = WMIService.ExecQuery(WMIQuery,,48)
    For Each drive in Drives
        LogDebug "Found that drive named "&VolumeName&" is known as "&drive.DeviceID
        GetDriveByVolumeName = drive.DeviceID
        Exit Function
    Next
    LogWarning "No Drive found whose name is " & VolumeName
End Function 'GetDriveByVolumeName --------------------------------------------


' -----------------------------------------------------------------------------
'        NAME: GetPhysicalDrive
' DESCRIPTION: Get the physical drive Id of a disk given its volume name
'     PARAM 1: Drive Letter of the disk to search physical name for
'      OUTPUT: Physical Id or empty string
' -----------------------------------------------------------------------------
Public Function GetPhysicalDrive(DriveLetter)
    WMIQuery = "ASSOCIATORS OF {Win32_LogicalDisk.DeviceID=""" & DriveLetter & """}" _
             & " WHERE AssocClass = Win32_LogicalDiskToPartition"
    Set Partitions = WMIService.ExecQuery(WMIQuery,,48)
    For Each part in Partitions
        PartID = part.DeviceID
    Next
    If PartID = "" Then
        TaskError "No partition found whose drice letter is " & DriveLetter
        Exit Function
    End If

    WMIQuery = "ASSOCIATORS OF {Win32_DiskPartition.DeviceID=""" & PartID & """}" _
             & " WHERE AssocClass = Win32_DiskDriveToDiskPartition"
   Set PhysicalDrives = WMIService.ExecQuery(WMIQuery,,48)
   For Each drive in PhysicalDrives
       LogDebug "Found that drive "&DriveLetter&" is known as "&drive.DeviceID
       GetPhysicalDrive = drive.DeviceID
       Exit Function
   Next
   LogWarning "No partition found whose drice letter is " & DriveLetter
End Function 'GetPhysicalDrive -------------------------------------------------
