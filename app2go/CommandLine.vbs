' -----------------------------------------------------------------------------
'         NAME: CommandLine.vbs
'  DESCRIPTION: Simple helpers to manipulate the command line
'      CREATED: 2014.04.23 / REVISION: 2014.07.30 - 16:02
' -----------------------------------------------------------------------------

' -----------------------------------------------------------------------------
'       NAME: AddCmdLine
' DESCIPTION: concatanate two parts of a command line
'    PARAM 1: first part of the command line
'    PARAM 2: second part of eth command line
'     RETURN: 1st part + second part
' -----------------------------------------------------------------------------
Public Function AddCmdLine(CLArgs, AddCLArgs)
    If CLArgs = "" Then
        AddCmdLine = AddCLArgs
    End If
    If AddCLArgs = "" Then
        AddCmdLine = CLArgs
    End If
    AddCmdLine = CLArgs & " " & AddCLArgs
End Function 'AddCmdLine -------------------------------------------------------


' -----------------------------------------------------------------------------
'       NAME: AddCmdLineOption
' DESCIPTION: Add a command line option only if not already present
'    PARAM 1: Array name of the option names to add. Ex: ("-p", "--profile")
'    PARAM 2: value of the option to add
'    PARAM 3: current command line
'     RETURN: optionName + option value + commandline
' -----------------------------------------------------------------------------
Public Function AddCmdLineOption(OptionName, OptionValue, CLArgs)
    newCmdLine = OptionName(0) & OptionValue
    For Each opt in OptionName
        If InStr(CLArgs, opt) > 0 Then
            LogDebug "Command Line Option "&opt&" is already set. Don't add it again"
            AddCmdLineOption = CLArgs
        End If
    Next
    AddCmdLineOption = AddCmdLine(newCmdLine, CLArgs)
End Function 'AddCmdLineOption -------------------------------------------------


' -----------------------------------------------------------------------------
'       NAME: SanitizeCmdLineArg
' DESCIPTION: Add additional " to take care of arguments with space that can
'             confuse WShell.Run
'    PARAM 1: argument to sanitize
'     RETURN: sanitized argument
'       TODO: take care also of additional cases than space according to:
'             http://blogs.msdn.com/b/twistylittlepassagesallalike/archive/2011/04/23/everyone-quotes-arguments-the-wrong-way.aspx?Redirected=true
' -----------------------------------------------------------------------------
Public Function SanitizeCmdLineArg(arg)
    If InStr(arg, " ") <> 0 Then
        SanitizeCmdLineArg = Chr(34) & Chr(34) & arg & Chr(34) & Chr(34)
    Else
        SanitizeCmdLineArg = arg
    End If
End Function 'SanitizeCmdLineArg -----------------------------------------------

