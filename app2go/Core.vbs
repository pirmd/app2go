' -----------------------------------------------------------------------------
'         NAME: core.vbs
'  DESCRIPTION: core app2go functions
'      CREATED: 2008.08.08 / REVISION: 2014.07.30 - 16:02
' -----------------------------------------------------------------------------

On Error Goto 0

' --- Bread an Butter objects and collections ---------------------------------
Set WshShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
Set WMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

'---+ Load the rest of App2go modules ------------------------------------------
include "app2go\Log.vbs"
include "app2go\LoadScript.vbs"
include "app2go\UsefulVar.vbs"
include "app2go\EnvAndVar.vbs"
include "app2go\Link.vbs"
include "app2go\Run.vbs"
include "app2go\RunService.vbs"
include "app2go\Registry.vbs"
include "app2go\ReplaceInFile.vbs"
include "app2go\RemovableDrive.vbs"
include "app2go\CommandLine.vbs"


' --- Tasks queues And Common resources ---------------------------------------
Set MyLauncher = New Launcher
Set CommonResources = New ShareResource

'Commandline arguments pass to the app2go script
CommandLineArguments = ""

'---+ Flags/Options -----------------------------------------------------------
FORCE_CLEAN = False


' ------------------------------------------------------------------------------
'         NAME: AddTaskXXX
'  DESCRIPTION: Add a launcher task for XXX 'run level'.
'      PARAM 1: task to run
'      PARAM 2: task arguments as Array()
' ------------------------------------------------------------------------------
Public Sub AddTaskStart(cmd, arg_array)
    MyLauncher.AddTaskStart cmd, arg_array
End Sub

Public Sub AddTaskStop(cmd, arg_array)
    MyLauncher.AddTaskStop cmd, arg_array
End Sub

Public Sub AddTaskOnExit(cmd, arg_array)
    MyLauncher.AddTaskOnExit cmd, arg_array
End Sub
' End AddTaskXXX --------------------------------------------------------------

' ------------------------------------------------------------------------------
'         NAME: TaskError
'  DESCRIPTION: Raise an Error during a task execution
'      PARAM 1: Error description
' ------------------------------------------------------------------------------
Public Sub TaskError(msg)
    MyLauncher.TaskError msg
End Sub ' End TaskError --------------------------------------------------------

' ------------------------------------------------------------------------------
'         NAME: Acquire/ReleaseResource
'  DESCRIPTION: Acquire/Release a shared resource
'      PARAM 1: array with resource ID
'       OUTPUT: True if acquisition/release is done for the first/last time
' ------------------------------------------------------------------------------
Public Function AcquireResource(resID)
    AcquireResource = CommonResources.Acquire(resID)
End Function

Public Function ReleaseResource(resID)
    ReleaseResource = CommonResources.Release(resID)
End Function
' End Acquire/ReleaseResource --------------------------------------------------

' ------------------------------------------------------------------------------
'         NAME: Set/GetResourceParam
'  DESCRIPTION: Set/Get a param of a shared resource
'      PARAM 1: array with resource ID
' ------------------------------------------------------------------------------
Public Sub SetResourceParam(resID, paramValue)
    CommonResources.SharedParam(resID) = paramValue
End Sub

Public Function GetResourceParam(resID)
    GetResourceParam = CommonResources.SharedParam(resID)
End Function
' End Set/getResourceParam -----------------------------------------------------


' ------------------------------------------------------------------------------
'         NAME: StartLauncher
'  DESCRIPTION: Start launcher
'      PARAM 1: Path to the script to start
'      PARAM 2: Additional arguments pass to the scripts
' ------------------------------------------------------------------------------
Public Sub StartLauncher(app2goScript, ScriptArguments)
    LogVerbose "Start "&app2goScript&" with arguments: "&ScriptArguments
    CommandLineArguments = ScriptArguments
    LoadScript app2goScript
    MyLauncher.StartLauncher()
End Sub 'StartLauncher ---------------------------------------------------------


' -----------------------------------------------------------------------------
'          CLASS: Launcher
'    DESCRIPTION: base class for a portable launcher which in fact
'                 aims at mimicing a try...finaly class
'                 I would love to get rid of it if someone has a better idea to
'                 ensure TodoClean/ToDoOnExit actions can be launched whatever happen
' PUBLIC METHODS:
'     . AddTask       -- Add a Task
'     . AddTaskClean  -- Add a Task to clean environement after tasks have been executed
'     . AddTaskOnExit -- Add a Task to perform before exiting launcher (and after cleaning)
'     . StartLauncher -- Execute the job in the queue.
'                        Take StopOnError flag as aparameter to determine if we
'                        stop or not the execution if an error appears
'     . StopLauncher  -- Stop the launcher and cleanly exit (execute the tasks
'                        in the TodoOnExit queue)
'     . Error         -- Handle errors during the execution of tasks
' TODO: Show in debug level the content of the queues
' -----------------------------------------------------------------------------
Class Launcher
    Private ToDoStart
    Private ToDoStop
    Private ToDoOnExit
    Private BreakOnError
    Private HasBeenStarted 'Flag that prevent CLass_terminate cleaning actions
                           'to occure if Launcher has not been started. Typically
                           'the case for shortcut2go.wsf that read a 2go script
                           'to get information without actually running it.
                           'Not a brillant move but it works

    Private Sub Class_initialize()
        Set ToDoStart = New TaskQueue
        Set ToDoStop = New TaskQueue
        Set ToDoOnExit = New TaskQueue
        BreakOnError = True
        HasBeenStart = False
    End Sub

    Public Sub StartLauncher()
        LogVerbose "Start launcher"
        HasBeenStart = True
        ToDoStart.start_FIFO(BreakOnError)
    End Sub

    Private Function build_task_cmd(cmd, arg_array)
        arg = Chr(34) & Join(arg_array, """,""") & Chr(34)
        build_task_cmd = Join(Array(cmd, arg), " ")
    End Function

    Public Sub AddTaskStart(cmd, arg_array)
        ToDoStart.Add(build_task_cmd(cmd&"_start", arg_array))
    End Sub

    Public Sub AddTaskOnExit(cmd, arg_array)
        ToDoOnExit.Add(build_task_cmd(cmd&"_exit", arg_array))
    End Sub

    Public Sub AddTaskStop(cmd, arg_array)
        ToDoStop.Add(build_task_cmd(cmd&"_stop", arg_array))
    End Sub

    Public Sub TaskError(Msg)
        If BreakOnError Then
            LogError Msg
            WScript.Quit(10)
        Else
            LogWarning Msg
            Err.Clear
        End If
    End Sub

    Private Sub Class_Terminate()
        If HasBeenStarted Then
            LogVerbose "Stopping launcher"
            BreakOnError = False 'try to execute as much tasks as possible on exit
            ToDoStop.start_LIFO(BreakOnError)
            ToDoOnExit.start_FIFO(BreakOnError)
            HasBeenSTarted = False
        End If
    End Sub
End Class 'Launcher -----------------------------------------------------------

' ------------------------------------------------------------------------------
'          CLASS: TaskQueue
'    DESCRIPTION: Poor-man task queue manager
' PUBLIC METHODS:
'     . Reset    -- clear queue
'     . Add      -- allow to add task to a queue
'     . Run_FIFO -- exec the tasks in the queue in FIFO mode
'     . Run_LIFO -- exec the tasks in the queue in LIFO mode
' -----------------------------------------------------------------------------
Class TaskQueue
  Private TaskList
  Private LastTaskIndex

  Private Sub Class_Initialize()
    Redim TaskList(20)
    LastTaskIndex = -1
  End Sub

  Public Sub Add(task)
    If LastTaskIndex=UBound(TaskList) then
      Redim Preserve TaskList(LastTaskIndex + 1)
    End If
    LastTaskIndex = LastTaskIndex + 1
    TaskList(LastTaskIndex) = task
  End Sub

 Private Sub execute_task(index)
     LogDebug "Execute task "&index&": "&TaskList(index)
     Execute TaskList(index)
 End Sub

 Public Sub start_LIFO(BreakOnError)
     If Not BreakOnError Then
         On Error Resume Next
     End If
     For task_index = LastTaskIndex To 0 Step -1
         execute_task(task_index)
     Next
     On Error Goto 0
 End Sub

 Public Sub start_FIFO(BreakOnError)
     If Not BreakOnError Then
         On Error Resume Next
     End If
     For task_index = 0 To LastTaskIndex
         execute_task(task_index)
     Next
     On Error Goto 0
 End Sub
End Class 'TaskQueue ----------------------------------------------------------

' -----------------------------------------------------------------------------
'          CLASS: ShareResource
'    DESCRIPTION: Poor-man semaphore-like to work with launcher share resources
'                 in a gentle manner (don't clean a resources used by someone else)
'     LIMITATION: Race conditions are not excluded when trying to read or update
'                 a shared environment variable but shoul dbe acceptable in our
'                 use case.
'                 hashcode use to compute environment var name according to
'                 resource name is weak but shoul dbe acceptable in our use case
' PUBLIC METHODS:
'     . Acquire  -- Acquire a resource. Return True if we are the first to
'                   acquire it, false otherwise
'     . Release  -- Release a resource. Return True if we are the last owner,
'                   False otherwise
'     . SetSharedParam -- define a parameter shared between launchers
'     . GetSharedParam -- retrieve a parameter shared between launchers
' TODO: Better FORCE_CLEAN Interface
' TODO: limit calls to resName
' -----------------------------------------------------------------------------
Class ShareResource
    Private semaphCollec
    Private paramCollec

    Private Sub Class_initialize()
        Set semaphCollec = WshShell.Environment("Volatile")
        Set paramCollec =  WshShell.Environment("Volatile")
    End Sub

    Private Property Get Semaphore(resID)
        varValue = semaphCollec(resName(resID))
        If varValue = "" Then 'First time we ask for this resource
            Semaphore = 0
        Else
            Semaphore = cint(varValue)
        End If
    End Property

    Private Property Let Semaphore(resID, varValue)
        varName = resName(resID)
        If varValue = 0 Then 'no more used
            LogDebug "Delete semaphore for shared resource "&resID(0)
            semaphCollec.Remove varName
            DelSharedParam(varName)
        Else
            semaphCollec(varName) = cstr(varValue)
        End If
    End Property

    Private Function resName(resID)
        resName = HashCode(Join(resID, "::"))
    End Function

    Private Function resNameParam(varName)
        resNameParam = varName&"_PARAM"
    End Function

    Private Function HashCode(myStr)
    ' F(n)=((127 & F(n-1) + string(n)) Mod 16908799
        hash = 0 : i =0
        while (i < Len(myStr))
            i = i + 1
            hash = ((127 * hash) + Asc(Mid(myStr,i,1))) Mod 16908799
        wend
        HashCode = hash
    End Function

    Public Function Acquire(resID)
        If FORCE_CLEAN Then
            LogDebug "Force resource acquisition for "& resID(0)
            Acquire = True
            Semaphore(resID) = 1
        Else
            varValue = Semaphore(resID)
            LogDebug "Acquire resource for "&resID(0)&" (shared with "&varValue&" launcher(s))"
            Semaphore(resID) = varValue+1
            Acquire = ( varValue = 0 )
        End If
    End Function

    Public Function Release(resID)
        If FORCE_CLEAN Then
            LogDebug "Force release of resource for "&resID(0)
            Semaphore(resID) = 0
            Release = True
        Else
            varValue = Semaphore(resID)
            LogDebug "Release resource for "&resID(0)&" (shared with "&varValue&" launcher(s))"
            Semaphore(resID) = varValue-1
            Release = (varValue = 1)
        End If
    End Function

    Public Property Let SharedParam(resID, ParamValue)
        varName = resNameParam(resID)
        paramCollec(VarName) = ParamValue
        LogDebug "Set shared parameter for "&resID(0)&" to "&ParamValue
    End Property

    Public Property Get SharedParam(resID)
        varName = resNameParam(resID)
        SharedParam = paramCollec(varName)
        LogDebug "Get shared parameter for "&resID(0)&" ("&SharedParam&")"
    End Property

    Private Sub DelSharedParam(varName)
        paramName = resNameParam(varName)
        If paramCollec(paramName) <> "" Then
            LogDebug "Delete shared parameter for "&resID(0)
            paramCollec.Remove paramName
        End If
    End Sub
End Class 'ShareResource ------------------------------------------------------


' ------------------------------------------------------------------------------
'        NAME: include
' DESCRIPTION: Load a vbs module
'     PARAM 1: Relative path to the module to load. Path is relatif to the main
'              script
' ------------------------------------------------------------------------------
Private Sub include(moduleRelPath)
    modulePath = FSO.GetParentFolderName(WScript.ScriptFullName)&"\"&moduleRelPath
    set module = FSO.OpenTextFile(modulePath)
    moduleCode = module.readAll()
    module.close
    executeGlobal moduleCode
End Sub 'include ---------------------------------------------------------------
