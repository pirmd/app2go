' -----------------------------------------------------------------------------
'         NAME: ReplaceInFile.vbs
'  DESCRIPTION: add/remove/delete from registry
'      CREATED: 2012.08.09 / REVISION: 2014.07.30 - 16:05
' -----------------------------------------------------------------------------

' -----------------------------------------------------------------------------
'        NAME: FixDriveLetter
' DESCRIPTION: Replace %LauncherDrive% with the drive letter from which the prog
'              is launched
'     PARAM 1: File to modify
' -----------------------------------------------------------------------------
Public Sub FixDriveLetter(FilePath)
    ReplaceInFile FilePath, "%LauncherDrive%", LauncherDrive
End Sub 'FixDriveLetter -------------------------------------------------------

' -----------------------------------------------------------------------------
'        NAME: FixUserFolders
' DESCRIPTION: Replace in a text file any reference to user folders by their
'              actual value. In that order:
'              - %DocFolder% by the folder name that contains the user documents
'              - %RootFolder% by the folder name that contains the App2go platform
'              - %LauncherDrive% by the drive letter that contains the launcher
'     PARAM 1: File to modify
' -----------------------------------------------------------------------------
Public Sub FixUserFolders(FilePath)
    If DocFolder <> "" Then
        ReplaceInFile FilePath, "%DocFolder%", DocFolder
    End If
    If RootFolder <> "" Then
        ReplaceInFile FilePath, "%RootFolder%", RootFolder
    End If
    ReplaceInFile FilePath, "%LauncherDrive%", LauncherDrive
End Sub 'FixUserFolders -------------------------------------------------------


' -----------------------------------------------------------------------------
'        NAME: ReplaceInFile
' DESCRIPTION: Replace a string in a text file
'     PARAM 1: File to modify
'     PARAM 2: Value to be replaced
'     PARAM 3: Replacement value
' -----------------------------------------------------------------------------
Public Sub ReplaceInFile(FilePath, ToReplace, ReplaceBy)
    AddTaskStart "ReplaceInFile", Array(FilePath, ToReplace, ReplaceBy)
End Sub

Private Function ReplaceInFile_resID(FilePath, ToReplace, ReplaceBy)
    ReplaceInFile_resID = Array("ReplaceInFile",FilePath, ToReplace, ReplaceBy)
End Function

Private Sub ReplaceInFile_start(FilePath, ToReplace, ReplaceBy)
    If AcquireResource(ReplaceInFile_resID(FilePath, ToReplace, ReplaceBy)) Then
        ReplaceStringInFile FilePath, ToReplace, ReplaceBy
        LogVerbose "Replace "&ToReplace&" by "&ReplaceBy&" in "&FilePath
    Else
        LogVerbose "Replacement of "&ToReplace&" by "&ReplaceBy&" in "&FilePath&" was done by another launcher"
    End If
    AddTaskStop "ReplaceInFile", Array(FilePath, ToReplace, ReplaceBy)
End Sub

Private Sub ReplaceInFile_stop(FilePath, ToReplace, ReplaceBy)
    If ReleaseResource(ReplaceInFile_resID(FilePath, ToReplace, ReplaceBy)) Then
        ReplaceStringInFile FilePath, ReplaceBy, ToReplace
        LogVerbose "Replace "&ReplaceBy&" by "&ToReplace&" in "&FilePath
    Else
        LogVerbose "File "&FilePath&" is used by another launcher. Don't restore it"
    End If
End Sub
' End Macro ReplaceInFile ------------------------------------------------------

Private Sub ReplaceStringInFile(FilePath, OldText, NewText)
    Const ForReading = 1
    Const ForWriting = 2
    LogDebug "Replace "&OldText&" by "&NewText&" in "&FilePath
  'Fail gracefully if file does not exists
    If Not FSO.FileExists(FilePath) Then
        LogWarning "File "&FilePath&" does not exist. Cannot replace "&OldText&" by "&NewText
        Exit Sub
    End If
  'Read and replace
    Set file = FSO.OpenTextFile(FilePath, ForReading)
    text = File.ReadAll
    file.Close
    newText = Replace(text, OldText, NewText)
   'Save new text
    Set file = FSO.OpenTextFile(FilePath, ForWriting)
    file.WriteLine newText
    file.Close
End Sub
