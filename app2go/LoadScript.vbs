' -----------------------------------------------------------------------------
'         NAME: app2go_LoadScript.vbs
'  DESCRIPTION: load an app2go script
'      CREATED: 2012.08.08 / REVISION: 2014.07.30 - 16:03
' -----------------------------------------------------------------------------
' TODO: merge with core.vbs 'include' ?

' -----------------------------------------------------------------------------
'        NAME: LoadScript
' DESCRIPTION: load an app2go script from a file
'     PARAM 1: path to script absolute or relative to current dir
' -----------------------------------------------------------------------------
Public Sub LoadScript(scriptPath)
    scriptPath = FSO.GetAbsolutePathName(ScriptPath)
    If FSO.FileExists(scriptPath) Then
        load_script scriptPath
    Else
        Die "Script "&scriptPath&" cannot be found."
    End If
End Sub 'LoadScript -----------------------------------------------------------

' -----------------------------------------------------------------------------
'        NAME: LoadScriptIfExists
' DESCRIPTION: load an app2go script from a file, doesn't die if file doesn't
'              exist
'     PARAM 1: path to script absolute or relative to current dir
' -----------------------------------------------------------------------------
Public Sub LoadScriptIfExists(scriptPath)
    scriptPath = FSO.GetAbsolutePathName(ScriptPath)
    If FSO.FileExists(scriptPath) Then
        load_script scriptPath
    Else
        LogWarning "Script "&scriptPath&" cannot be found. Pass loading"
    End If
End Sub 'LoadScript -----------------------------------------------------------


Private Sub load_script(scriptPath)
    LogDebug "Load script "&scriptPath
  'Change the current directory to allow 'LoadScript' work with relative
  'path from within the script we are including
    currDir = WshShell.CurrentDirectory
    WshShell.CurrentDirectory = FSO.GetParentFolderName(scriptPath)
  'Source script content
    set script = FSO.OpenTextFile(scriptPath)
    scriptCode = script.readAll()
    scriptCode = WshShell.ExpandEnvironmentStrings(scriptCode)
    script.close
    executeGlobal scriptCode
  'Restore current directory
    WshShell.CurrentDirectory = currDir
end sub
