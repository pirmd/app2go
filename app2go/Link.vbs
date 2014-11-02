' -----------------------------------------------------------------------------
'         NAME: Link.vbs
'  DESCRIPTION: create links
'      CREATED: 2012.08.07 / REVISION: 2014.07.30 - 16:02
' -----------------------------------------------------------------------------
' TODO: [link_start] behavior if target already exists


' -----------------------------------------------------------------------------
'        NAME: Link
' DESCRIPTION: Link source to dest:
'             - If source is a folder, create a NTFS junction,
'             - If a source is a file, create a shortcut
'             - If dest is a drive letter, create a subst mount point
'     DEPENDS: junction.exe from sysinternal. junction.exe should be in PATH
'     PARAM 1: source, either a folder or a file
'     PARAM 2: destination.
' -----------------------------------------------------------------------------
Public Sub Link(source, target)
    AddTaskStart "link", Array(source, target)
End Sub

Private Function link_resID(source, target)
    link_resID = Array("link", source, target)
End Function

Private Sub link_start(source, target)
    If AcquireResource(link_resID(source, target)) Then
        If Right(target,1) = ":" Then
            LogVerbose "Create subst from "&source&" to "&target
            WshShell.Run "subst.exe """&target&""" """&source&"\""", RUN_HIDDEN, WAIT_FOR_ME

        ElseIf FSO.FolderExists(source) Then
            LogVerbose "Create junction from "&source&" to "&target
            WshShell.Run "junction.exe """&target&""" """&source&"""", RUN_HIDDEN, WAIT_FOR_ME

        ElseIf FSO.FileExists(source) Then
            LogVerbose "Create shortcut from "&source&" to "&target
            target = target & ".lnk"
            CreateShortcut target, source, ""

        Else
            TaskError "Link "&source&" to "&target&": "&source&" does not exist."
            ReleaseResource(link_resID(source, target))
            Exit Sub
        End If
    Else
        LogVerbose target&" was linked to "&source&" by another launcher"
    End If
    AddTaskStop "link", Array(source, target)
End Sub

Private Sub link_stop(source, target)
    If ReleaseResource(link_resID(source, target)) Then
        If Right(target,1) = ":" Then
            LogVerbose "Delete subst from "&source&" to "&target
            WshShell.Run "subst.exe """&target&""" /D", RUN_HIDDEN, WAIT_FOR_ME

        ElseIf FSO.FolderExists(target) Then
            LogVerbose "Delete junction from "&source&" to "&target
            WshShell.Run "junction.exe -d """&target&"""", RUN_HIDDEN, WAIT_FOR_ME

        ElseIf FSO.FileExists(target) Then
            LogVerbose "Delete shortcut from "&source&" to "&target
            FSO.GetFile(target).Delete(True)
        End If
    Else
        LogVerbose target&" is used by another launcher. Don't remove it"
    End If
End Sub
' End Macro Link --------------------------------------------------------------


' -----------------------------------------------------------------------------
'        NAME: Shortcut
' DESCRIPTION: Create a shortcut
'     PARAM 1: source, either a folder or a file
'     PARAM 2: destination.
'     PARAM 3: additional arguments to pass to the prog
' -----------------------------------------------------------------------------
Public Sub MakeShortcut(source, target, arguments)
    AddTaskStart "MakeShortcut", Array(source, target, arguments)
End Sub

Private Function MakeShortcut_resID(source, target)
    MakeShortcut_resID = Array("MakeShortcut", source, target)
End Function

Private Sub MakeShortcut_start(source, target, arguments)
    If AcquireResource(MakeShortcut_resID(source,target)) Then
        LogVerbose "Create shortcut from "&source&" to "&target
        CreateShortcut target&".lnk", source, arguments
    Else
        LogVerbose target&" was linked to "&source&" by another launcher"
    End If
    AddTaskStop "MakeShortcut", Array(source, target)
End Sub

Private Sub MakeShortcut_stop(source, target)
    If ReleaseResource(MakeShortcut_resID(source,target)) Then
        LogVerbose "Delete shortcut from "&target&" to "&source
        FSO.GetFile(target&".lnk").Delete(True)
    Else
        LogVerbose target&" is used by another launcher. Don't remove it"
    End If
End Sub
' End Macro Shortcut -----------------------------------------------------------


' -----------------------------------------------------------------------------
'        NAME: AddToDesktop
' DESCRIPTION: Create a link available from the Desktop
'     PARAM 1: source, either a folder or a file
'     PARAM 2: additional arguments to pass to the prog
'     PARAM 3: shortcut name
' -----------------------------------------------------------------------------
Public Sub AddToDesktop(source, arguments, targetName)
    target = WshShell.SpecialFolders("Desktop") & "\" & targetName
    MakeShortcut source, target, arguments
End Sub 'AddToDesktop ---------------------------------------------------------

' -----------------------------------------------------------------------------
'        NAME: AddToSendTo
' DESCRIPTION: Create a link available from the 'Send To' right click menu
'     PARAM 1: source, either a folder or a file
'     PARAM 2: additional arguments to pass to the prog
'     PARAM 3: shortcut name
' -----------------------------------------------------------------------------
Public Sub AddToSendTo(source, arguments, targetName)
    target = WshShell.SpecialFolders("SendTo") & "\" & targetName
    MakeShortcut source, target, arguments
End Sub 'AddToSendTo ----------------------------------------------------------



Private Sub CreateShortcut(source, target, arguments)
    LogDebug "Create a shortcut from "&source&" to "&target
    set shortcut = WshShell.CreateShortcut(source)
    shortcut.TargetPath = target
    shortcut.Arguments = arguments
    shortcut.save
End Sub
