' -----------------------------------------------------------------------------
'         NAME: Registry.vbs
'  DESCRIPTION: add/remove/delete from registry
'      CREATED: 2012.08.09 / REVISION: 2014.10.30 - 13:39
' -----------------------------------------------------------------------------
' TODO: Switch to using WMI instead of WshShell functions to access registry

' -----------------------------------------------------------------------------
'        NAME: MergeHiveFile
' DESCRIPTION: Merge a Hive reg file. A hive is a logical group of keys, subkeys,
'              and values in the registry that has a set of supporting files
'              containing backups of its data.
'     PARAM 1: Key to be set or modified
'     PARAM 2: Value of the key
' TODO: get RegParentKey from HiveFilePath
' -----------------------------------------------------------------------------
Public Sub MergeHiveFile(HiveFilePath, RegParentKey)
    AddTaskStart "MergeHiveFile", Array(HivefilePath, RegParentKey)
End Sub

Private Function MergeHiveFile_resID(HiveFilePath, RegParentKey)
    MergeHiveFile_resID = Array("MergeHiveFile", HiveFilePath, RegParentKey)
End Function

Private Sub MergeHiveFile_start(HivefilePath, RegParentKey)
    resID = MergeHiveFile_resID(HivefilePath, RegParentKey)
    If AcquireResource(resID) Then
        restoreHiveFileName = FSO.GetBaseName(HiveFilePath)&"-BACKUP.reg"
        restoreHiveFilePath = FSO.GetParentFolderName(RegFilePath)&"\"&restoreHiveFileName
        WshShell.Run "regedit.exe /E /S "&restoreHiveFilePath&" "&RegParentPath, RUN_HIDDEN, WAIT_FOR_ME
        SetResourceParam resID, restoreHiveFilePath
        WshShell.RegDelete RegParentKey
        WshShell.Run "regedit.exe /S "&HiveFilePath, RUN_HIDDEN, WAIT_FOR_ME
        LogVerbose "Merge "&HiveFilePath&" into the registry"
    Else
        LogVerbose HiveFilePath&" was already merged by another launcher"
        restoreHiveFilePath = GetResourceParam(resID)
    End If
    AddTaskStop "MergeHiveFile", Array(HiveFilePath, restoreHiveFilePath, RegParentKey)
End Sub

Private Sub MergeHiveFile_stop(HivefilePath, restoreHiveFilePath, RegParentKey)
    If ReleaseResource(MergeHiveFile_resID(HiveFilePath, RegParentKey)) Then
        LogVerbose "Save modification of "&RegParentKey&" to "&HiveFilePath
        WshShell.Run "regedit.exe /E /S "&HiveFilePath&" "&RegParentPath, RUN_HIDDEN, WAIT_FOR_ME
        LogVerbose "Restore "&RegParentKey&" to it's original state"
        WshShell.RegDelete RegParentKey
        WshShell.Run "regedit.exe /S "&restoreHiveFilePath, RUN_HIDDE, WAIT_FOR_ME
        FSO.DeleteFile(restoreHiveFilePath)
    Else
        LogVerbose HiveFilePath&" is used by another launcher. Don't restore it"
    End If
End Sub
' End Macro MergeHiveFile ------------------------------------------------------

' -----------------------------------------------------------------------------
'        NAME: SetRegKey
' DESCRIPTION: Define a registry key value and restore it afterwards
'     PARAM 1: Key to be set or modified
'     PARAM 2: Value of the key
' -----------------------------------------------------------------------------
Public Sub SetRegKey(KeyName, KeyValue)
    AddTaskStart "SetRegKey", Array(KeyName, KeyValue)
End Sub

Private Function SetRegKey_resID(KeyName, KeyValue)
    SetRegKey_resID = Array("SetRegKey", KeyName, KeyValue)
End Function

Private Sub SetRegKey_start(KeyName, KeyValue)
    resID = SetRegKey_resID(KeyName, KeyValue)
    If AcquireResource(resID) Then
        restoreValue = ReadRegKey(KeyName)
        SetResourceParam resID, restoreValue
        WshShell.RegWrite KeyName, KeyValue
        LogVerbose "Set registry key "&KeyName&" to "&ReadRegKey(KeyName)
    Else
        LogVerbose "Registry key "&KeyName&" was defined by another launcher to "&ReadRegKey(KeyName)
        restoreValue = GetResourceParam(resID)
    End If
    AddTaskStop "SetRegKey", Array(KeyName, KeyValue, restoreValue)
End Sub

Private Sub SetRegKey_stop(KeyName, KeyValue, restoreValue)
    If ReleaseResource(SetRegKey_resID(KeyName, KeyValue)) Then
        If restoreValue="" Then
            LogVerbose "Remove registry key "&KeyName
            WshShell.RegDelete KeyName
        Else
            LogVerbose "Restore registry key "&KeyName&" to "&restoreValue
            WshShell.RegWrite KeyName, restoreValue
        End If
    Else
        LogVerbose "Registry key "&KeyName&" is used by another launcher. Don't restore it"
    End If
End Sub
' End Macro SetRegKey ---------------------------------------------------------


'Isn't it a beautiful piece of code, just to read a value in the registry ?
'I *LOVE* microsoft way of doing things...
'I really should have miss something
Private Function GetRegKeyDontExistErrDesc()
    Const FakeKey = "HKEY_DONT_EXIST\"
    On Error Resume Next
    WshShell.RegRead FakeKey
    GetRegKeyDontExistErrDesc = Replace(Err.Description, FakeKey, "")
    Err.Clear
    On Error Goto 0
End Function
RegKeyDontExists = GetRegKeyDontExistErrDesc()

Private Function ReadRegKey(KeyToRead)
    KeyName = KeyToRead
    If Not Right(KeyName, 1)="\" Then
        KeyName = KeyName & "\"
    End If
    On Error Resume Next
    RegReadKey = WshShell.RegRead(KeyName)
    If Replace(Err.Description, KeyName, "")=RegKeyDontExists Then
        LogDebug KeyName&" does not exist"
        RegReadKey = ""
        Err.Clear
    End If
    On Error Goto 0
End Function
