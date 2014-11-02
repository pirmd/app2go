' -----------------------------------------------------------------------------
'         NAME: app2go_Run.vbs
'  DESCRIPTION: Run commands/apps
'      CREATED: 2008.08.08 / REVISION: 2014.10.30 - 13:51
' -----------------------------------------------------------------------------

'---+ Nice Constants ----------------------------------------------------------
Const WAIT_FOR_ME = "True"
Const DONT_WAIT_FOR_ME = "False"
Const RUN_HIDDEN = 0
Const RUN_NORMAL = 1
Const RUN_MINIMIZED = 2
Const RUN_MAXIMIZED = 3

' -----------------------------------------------------------------------------
'          NAME: Run
'   DESCRIPTION: Execute a command
'       PARAM 1: program to execute
'       PARAM 2: program arguments
'    PARAM 3, 4: same than WshShell.Run WindowStyle and WaitOnReturn
'   TODO: use only spawn_proc instead of WshRun
'   TODO: get directly a process object instead of a pID for run_stop
'   TODO: share more code with run_exit and run_start
' -----------------------------------------------------------------------------
Public Sub Run(cmd, arg, WindowStyle, WaitOnReturn)
    AddTaskStart "run", Array(cmd, arg, WindowStyle, WaitOnReturn)
End Sub

Private Sub run_start(cmd, arg, WindowStyle, WaitOnReturn)
    cmdLine = build_command_line(cmd, arg)
    If WaitOnReturn Then
        LogVerbose "Run and wait for "&cmdLine
        exitCode = WshShell.Run(cmdLine, WindowStyle, WaitOnReturn)
        If exitCode <> 0 Then
            TaskError cmd&" failed to run with exit code "&exitCode
        Else
            LogDebug "Successfully run "&cmdLine
        End If
    Else
        out = spawn_proc(cmdLine, WindowStyle)
        If out(0) <> 0 Then 'exitcode
            TaskError "Fail to spawn "&cmdLine& "with exit code "&out(0)
        Else
            LogDebug "Successfully spawn "&cmdline
            AddTaskStop "run", Array(out(1))
        End If
    End If
End Sub

Private Sub run_stop(pid)
   WMIQuery = "SELECT * FROM Win32_Process WHERE ProcessId="""&pid&""""
    Set RunningProcesses = WMIService.ExecQuery(WMIQuery,,48)
    For Each proc in RunningProcesses
        LogVerbose "Killing: "&proc.CommandLine
        proc.Terminate
    Next
End Sub
'End Macro Run -----------------------------------------------------------------


' -----------------------------------------------------------------------------
'          NAME: RunOnExit
'   DESCRIPTION: Execute a command only when the launcher exit
'       PARAM 1: program to execute
'       PARAM 2: program arguments
'    PARAM 3, 4: same than WshShell.Run WindowStyle and WaitOnReturn
' -----------------------------------------------------------------------------
Public Sub RunOnExit(cmd, arg, WindowStyle, WaitOnReturn)
    AddTaskOnExit "Run", Array(cmd, arg, WindowStyle, WaitOnReturn)
End Sub

Private Sub Run_exit(cmd, arg, WindowStyle, WaitOnReturn)
    cmdLine = build_command_line(cmd, arg)
    If WaitOnReturn Then
        LogVerbose "Run and wait for "&cmdLine
        exitCode = WshShell.Run(cmdLine, WindowStyle, WaitOnReturn)
        If exitCode <> 0 Then
            TaskError cmd&" failed to run with exit code "&exitCode
        Else
            LogDebug "Successfully run "&cmdLine
        End If
    Else
        out = spawn_proc(cmdLine, WindowStyle)
        If out(0) <> 0 Then 'exitcode
            TaskError "Fail to spawn "&cmdLine& "with exit code "&out(0)
        Else
            LogDebug "Successfully spawn "&cmdline
        End If
    End If
End Sub
'End Macro RunOnExit -----------------------------------------------------------


Private Function build_command_line(cmd, arg)
    If arg <> "" Then
        build_command_line =  Chr(34) & cmd & Chr(34) & " " & arg
    Else
        build_command_line = Chr(34) & cmd & Chr(34)
    End If
End Function


Private Function spawn_proc(cmd, WindowStyle)
    Set ProcConfig = WMIService.Get("Win32_ProcessStartup").SpawnInstance_
    ProcConfig.ShowWindow = WindowStyle
    exitCode = WMIService.Get("Win32_Process").Create(cmd, Null, ProcConfig, ProcID)
    spawn_proc = Array(exitCode, ProcID)
End Function
