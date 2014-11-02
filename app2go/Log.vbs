' -----------------------------------------------------------------------------
'         NAME: Log.vbs
'  DESCRIPTION: vbs library that provide basic log facilities
'      CREATED: 2012.08.07 / REVISION: 2014.07.30 - 16:03
' -----------------------------------------------------------------------------
' TODO: Log to file


'---+ Nice Constants ----------------------------------------------------------
Const LOG_CRITICAL = 0
Const LOG_ERROR = 10
Const LOG_WARNING = 20
Const LOG_INFO = 30
Const LOG_VERBOSE = 40
Const LOG_DEBUG = 50

'---+ Easy Access to Log facilities from outside -------------------------------
Set MyLogger = New Logger

' ----------------------------------------------------------------------------
'        NAME: LogXXXX
' DESCRIPTION: Send a log message of level XXXX
' ----------------------------------------------------------------------------
Public Sub LogDebug(msg)
    MyLogger.debug msg
End Sub

Public Sub LogError(msg)
    MyLogger.error msg
End Sub

Public Sub LogVerbose(msg)
    MyLogger.verbose msg
End Sub

Public Sub LogInfo(msg)
    MyLogger.info msg
End Sub

Public Sub LogWarning(msg)
    MyLogger.warning msg
End Sub
' End Of functions LogXXXXX ---------------------------------------------------

' ----------------------------------------------------------------------------
'        NAME: Die
' DESCRIPTION: Quit program immediatly with an error
'     PARAM 1: error message
' ----------------------------------------------------------------------------
Private Sub Die(Msg)
    LogError Msg
    WScript.Quit(-1)
End Sub 'Die -----------------------------------------------------------------


' ------------------------------------------------------------------------------
'        NAME: ShowXXXXMsg
' DESCRIPTION: Show messages up to level XXXX
' ------------------------------------------------------------------------------
Public Sub ShowDebugMsg()
    MyLogger.LogLvl = LOG_DEBUG
End Sub

Public Sub ShowErrorMsg()
    MyLogger.LogLvl = LOG_ERROR
End Sub

Public Sub ShowVerboseMsg()
    MyLogger.LogLvl = LOG_VERBOSE
End Sub

Public Sub ShowInfoMsg()
    MyLogger.LogLvl = LOG_INFO
End Sub

Public Sub ShowWarningMsg()
    MyLogger.LogLvl = LOG_WARNING
End Sub
'End of functions ShowXXXXMsg --------------------------------------------------


' -----------------------------------------------------------------------------
'          CLASS: Logger
'    DESCRIPTION: Basic log set of functions. Allow at the moment to display
'                 messages to the console or as msgbox
' PUBLIC METHODS:
'     . critical, error, waring, info, debug  -- output log message of the
'                                                corresponding level
' -----------------------------------------------------------------------------
Class Logger
    Private LogLevel
    Public LogToUser

    Private Sub Class_Initialize()
        LogLevel = LOG_ERROR
        LogToUser = True
    End Sub

    Public Property Let LogLvl(lvl)
        'We swicth to console display if we are going to issue a lot of
        'message for better readability
        If (lvl >= LOG_VERBOSE) AND LogToUser Then
            ShowLogConsole()
        End If
        LogLevel = lvl
    End Property

    Private Sub ShowLogConsole()
        'TODO: is there a better way to send output to console?
        If Instr(WScriptInterpreter, "wscript.exe")>0 Then
            WshShell.Run "%COMSPEC% /c cscript.exe //nologo "&WScriptCommandLine()&" & pause"
            WScript.Quit(0)
        End If
    End Sub

    Private Sub LogToConsole(msg)
        If LogToUser Then
            WScript.Echo msg
        End If
    End Sub

    Private Sub message(lvl, msg)
        If lvl <= LogLevel Then
            LogToConsole msg
        End If
    End Sub

    Public Sub debug(msg)
        message LOG_DEBUG, "DEBUG    -- "&msg
    End Sub

    Public Sub verbose(msg)
        message LOG_VERBOSE, "VERBOSE  -- "&msg
    End Sub

    Public Sub info(msg)
        message LOG_INFO, "INFO     -- "&msg
    End Sub

    Public Sub warning(msg)
        message LOG_WARNING, "WARNING  -- "&msg
    End Sub

    Public Sub error(msg)
        message LOG_ERROR, "ERROR    -- "&msg
    End Sub

    Public Sub critical(msg)
        message LOG_CRITICAL, "CRITICAL -- "&msg
    End Sub
End Class
'End Class Logger ------------------------------------------------------------

Private Function WScriptInterpreter()
    WScriptInterpreter = LCase(WScript.FullName)
End Function

Private Function WScriptCommandLine()
    cmdline = Chr(34) & WScript.ScriptFullName & Chr(34)
    For Each arg in WScript.Arguments
        cmdline = cmdline & " " & arg
    Next
    WScriptCommandLine = cmdline
End Function

