' -----------------------------------------------------------------------------
'         NAME: EnvAndVar.vbs
'  DESCRIPTION: modify environment and PATH
'      CREATED: 2012.08.08 / REVISION: 2014.07.30 - 16:02
' -----------------------------------------------------------------------------


Set EnvVar = WshShell.Environment("Process")


' -----------------------------------------------------------------------------
'       NAME: SetEnv
' DESCIPRION: define an environment variable
'    PARAM 1: name of the variable
'    PARAM 2: value
' -----------------------------------------------------------------------------
Public Sub SetEnv(VarName, VarValue)
    AddTaskStart "SetEnv", Array(VarName,VarValue)
End Sub

Private Sub SetEnv_start(VarName, VarValue)
    EnvVar(VarName) = VarValue
    LogVerbose "Set environment variable "&VarName&"="&EnvVar(VarName)
End Sub
'End of Macro SetEnv ----------------------------------------------------------


' -----------------------------------------------------------------------------
'       NAME: AddToPATH
' DESCIPRION: Add a path to PATH
'    PARAM 1: path to add
' -----------------------------------------------------------------------------
Public Sub AddToPATH(path)
    AddTaskStart "AddToPATH", Array(path)
End Sub

Private Sub AddToPATH_start(path)
    If Instr(WshShell.ExpandEnvironmentStrings("%PATH%"), path) = 0 Then
        If EnvVar("PATH")="" Then
            EnvVar("PATH") = path
        Else
            EnvVar("PATH") = path&";"&EnvVar("PATH")
        End If
        LogVerbose "Add "&path&" to PATH -> PATH="&EnvVar("PATH")
    Else
        LogVerbose "Path "&path&" is already in PATH"
    End If
End Sub
'End Macro AddToPATH -----------------------------------------------------------


' -----------------------------------------------------------------------------
'       NAME: DefaultTo
' DESCIPRION: set a default value to a variable
'    PARAM 1: name of the variable
'    PARAM 2: default value
' -----------------------------------------------------------------------------
Public Sub DefaultTo(VarName, DefaultValue)
    Execute "varvalue="&VarName
    If varvalue = "" Then
        Execute varname & "=" & Chr(34) & defaultvalue & Chr(34)
        LogDebug "Set default value for variable "&VarName&" to "&DefaultValue
    Else
        LogDebug VarName&" is already set. Do not apply default value"
    End If
End Sub 'DefaultTo ------------------------------------------------------------

