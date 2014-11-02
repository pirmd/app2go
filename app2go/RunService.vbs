' -----------------------------------------------------------------------------
'         NAME: RunService.vbs
'  DESCRIPTION: run a service shared between several launchers
'      CREATED: 2012.08.08 / REVISION: 2014.07.30 - 16:06
' -----------------------------------------------------------------------------



' -----------------------------------------------------------------------------
'          NAME: RunService
'   DESCRIPTION: Execute a command as a shared 'service' for all lauchers
'       PARAM 1: program to execute
'       PARAM 2: program arguments
'       PARAM 3: same than WshShell.Run WindowStyle
' -----------------------------------------------------------------------------
Public Sub RunService(cmd, arg, WindowStyle)
    AddTaskStart "RunService", Array(cmd, arg, WindowStyle)
End Sub

Private Function RunService_resID(cmd, arg)
    RunService_resID = Array("RunService", cmd, arg)
End Function

Private Sub RunService_start(cmd, arg, WindowStyle)
    cmdLine = build_command_line(cmd, arg)
    resID = RunService_resID(cmd, arg)
    If AcquireResource(resID) Then
        If is_running(cmd, arg) > 0 Then
            LogVerbose cmd&" is already running but no launcher seems to reclaim it"
            ReleaseResource resID
        Else
            out = spawn_proc(cmdLine, WindowStyle)
            exitCode = out(0)
            pid = out(1)
            SetResourceParam resID, pid
            AddTaskOnExit "RunService", Array(cmdLine, pid)
            If exitCode <> 0 Then
                TaskError cmd&" failed to run with exit code "&exitCode
            Else
                LogVerbose "Successfully run service "&cmd
            End If
        End If
    Else
        servicePID = GetResourceParam(resID)
        AddTaskStop "RunService", Array(cmd, arg, ServicePID)
        LogVerbose cmd&" is already launched as a shared service (pid:"&ServicePID&")"
    End If
End Sub

Private Sub RunService_stop(cmd, arg, pid)
    cmdLine = build_command_line(cmd, arg)
    If ReleaseResource(RunService_resID(cmd, arg)) Then
        Run_stop pid
    Else
        LogVerbose cmd&" is used by another launcher. Don't kill it"
    End If
End Sub
' End Macro RunService --------------------------------------------------------


Private Function is_running(cmd, arg)
    'TODO: Look for program full path (use ExecutablePath instead of Name)?
    'TODO: Check if it's working with cmd with space
    ProcName = FSO.GetFileName(cmd) 'TODO: sould we use quote?
    CmdLine = "%"&join_str(ProcName, arg, "%") 'in WQL format for LIKE operator
    LogDebug "Check how many process "&cmd&" with arg "&arg&" are running"
    WMIQuery = "SELECT * FROM Win32_Process WHERE Name="""&ProcName&"""" _
             & " AND CommandLine LIKE """&CmdLine&""""
    Set RunningProcesses = WMIService.ExecQuery(WMIQuery,,48)
    'TODO: why RunningProcesses.Count doesn't work?
    For Each proc in RunningProcesses
        LogDebug "Running: "&proc.CommandLine
        procCount = procCount + 1
    Next
    is_running = (procCount > 0)
    LogDebug cmd&" is running : "&procCount&" times"
End Function
