VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Engine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Created by Robin G Brown, 21st July 1997
'-----------------------------------------------------------------------------
'   This class module contains code for _
    a simple script handling engine
'   Use sce as a prefix for this class
'-----------------------------------------------------------------------------
'   Script File Information
'
'   Script files must be in the following format(s):
'
'   [ACTION]|[IMPORT]|[EXPORT]|[PURGE]
'   server=SERVER_APPLICATION.SERVER_OBJECT
'   source=SOURCE_NAME
'   destination=DESTINATION_NAME
'   options=OPTIONS_ARE_SERVER_SPECIFIC_AND_OPTIONAL
'       note that there must be at least server and source data
'       and that these must follow the IServerScriptAbs interface
'
'   [SCHEDULE]
'   schedule=DATE_AND_TIME_TO_RUN
'       Note that the schedule action will cause the script to stop execution
'       at that point and it will be re-executed ENTIRELY at a later date/time
'       Thus the Schedule action is best used at the start of a script
'
'   [LAUNCH]
'   program=PROGRAM_COMMAND_LINE
'
'   [SLEEP]
'   sleep=TIME_TO_CONTINUE
'
'   [Log]
'   logfile=LOG_FILE_NAME
'
'   [NOLOG]
'
'   [FILE]
'   source=SOURCE_FILE_AND_OR_DIRECTORY
'   destination=OPTIONAL_DESTINATION_FILE_AND_OR_DIRECTORY
'   options=copy|move|delete|rename
'       only one of the above may be chosen
'
'   [DIRECTORY]
'   source=DIRECTORY
'   options=create|delete|rename
'       only one of the above may be chosen
'
'   [SCRIPT]
'   script=SCRIPT_FILE_TO_RUN
'
'   [END]
'
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conClass = "Engine"
'Private Types
Private Type EngineAction
    'Use act as a prefix
    ServerName As String
    SourceName As String
    DestinationName As String
    OptionList As String
End Type
'Private internal objects
Private objServer As Object
Private alaSleeper As Sleeper
Private WithEvents alaScheduler As Alarm
Attribute alaScheduler.VB_VarHelpID = -1
'Private properties
Private lngReturn As Long
Private strScriptFile As String
Private strLogFilename As String
Private strCurrentScriptLine As String
Private strLastError As String
Private lngLogFileNumber As Long
Private lngScriptFileNumber As Long
Private dteScheduledTime As Date
Private booLoggingFlag As Boolean
Private booAlarmWaiting As Boolean
Private booActionLine As Boolean
'Internal Error Codes
Private Const conNoError = "No Error"
Private Const conUndefinedError = "Undefined"
Private Const conLoggingActiveError = "Logging already active"
Private Const conIncorrectScheduleError = "Incorrect Schedule"
Private Const conMissingParametersError = "Missing Parameters"

Private Sub alaScheduler_AlarmRing()
'Created by Robin G Brown, 22/7/97
'This is the return event from the Alarm object
'Sub Declares
    'Error Trap
    On Error Resume Next
    'Open the script and run it
    WriteScriptEngineLog "Alarm rung"
    Call ProcessScript
    'Set booAlarmWaiting after the script has run so that it ignores the Schedule action
    booAlarmWaiting = False
    Set alaScheduler = Nothing
    If Err.Number <> 0 Then
        'There was an error, probably when running the script, _
            so put up a message for the user informing them of the problem
        MsgBox "There was an error (Error: " & Err.Number & ", '" & Err.Description & "') while running a scheduled script, please retry.", vbOKOnly + vbExclamation, "Script Engine"
    End If
End Sub

Private Sub Class_Initialize()
'Created by Robin G Brown, 21/7/97
'Default behaviour
'Sub Declares
    'Error Trap
    On Error Resume Next
    booAlarmWaiting = False
    strScriptFile = ""
    Set objServer = Nothing
    strLastError = conNoError
End Sub

Public Sub RunText(ByVal strScriptText As String)
'Created by Robin G Brown, 21/7/97
'Process the given text by saving it to a temporary file and running that
'Sub Declares
Const conSub = "RunText"
Dim strTempFile As String
Dim lngTempFileNumber As Long
    'Error Trap
    On Error GoTo RunText_ErrorHandler
    'Create a temporary file
    strScriptFile = App.Path & "\~SCRIPT.TXT"
    lngTempFileNumber = FreeFile()
    Open strScriptFile For Output As #lngTempFileNumber
    Print #lngTempFileNumber, strScriptText
    Close #lngTempFileNumber
    'Process it
    Call ProcessScript
    'and destroy it
    Kill strScriptFile
    strScriptFile = ""
Exit Sub
RunText_ErrorHandler:
    Select Case Err.Number
    Case Else
        WriteScriptEngineLog "Unexpected error : " & Err.Description & ", " & conClass & " : " & conSub
        Err.Raise Err.Number, conClass, Err.Source & "," & Err.Description, Err.HelpFile, Err.HelpContext
    End Select
End Sub

Public Sub RunFile(ByVal strScriptFileName As String)
'Created by Robin G Brown, 21/7/97
'Set the script file and process it
'Sub Declares
Const conSub = "RunFile"
    'Error Trap
    On Error GoTo RunFile_ErrorHandler
    strScriptFile = strScriptFileName
    Call ProcessScript
Exit Sub
RunFile_ErrorHandler:
    Select Case Err.Number
    Case Else
        WriteScriptEngineLog "Unexpected error : " & Err.Description & ", " & conClass & " : " & conSub
        Err.Raise Err.Number, conClass, Err.Source & "," & Err.Description, Err.HelpFile, Err.HelpContext
    End Select
End Sub

Private Sub ProcessScript()
'Created by Robin G Brown, 10/7/97
'Open the Script file, read the 'Actions', set the variables and run the server
'Sub Declares
Const conSub = "ProcessScript"
Dim strActionType As String
    'Error Trap
    On Error GoTo ProcessScript_ErrorHandler
    'Open the script file
    lngScriptFileNumber = FreeFile()
    Open strScriptFile For Input As #lngScriptFileNumber
    WriteScriptEngineLog "Script file '" & strScriptFile & "' opened"
    'Read the script file until we find an action to carry out
    Do While Not EOF(lngScriptFileNumber)
        Line Input #lngScriptFileNumber, strCurrentScriptLine
        If InStr(strCurrentScriptLine, "[") = 1 And InStr(strCurrentScriptLine, "]") > 0 Then
            Exit Do
        End If
    Loop
    'Now process actions until we run out or are told to stop
    Do While Not EOF(lngScriptFileNumber)
        'What sort of line is it?
        If InStr(strCurrentScriptLine, "[") = 1 And InStr(strCurrentScriptLine, "]") > 0 Then
            'Action definition line, strActionType allows for several different actions
            strActionType = Mid$(strCurrentScriptLine, 2, Len(strCurrentScriptLine) - 2)
            WriteScriptEngineLog "Action : " & strActionType
            'Note that the called action procedures read the next line, _
                if no action is called the next line MUST be read _
                or there will be an infinite loop!
            Select Case strActionType
            Case "SCHEDULE"
                If booAlarmWaiting = True Then
                    'If there is an alarm waiting then ignore the Schedule action
                     WriteScriptEngineLog "Alarm is set, ignoring Schedule action"
                     Line Input #lngScriptFileNumber, strCurrentScriptLine
                Else
                    'Exit this function setting up an alarm to continue later!
                    Call ScheduleAction
                    If strLastError <> "ScheduleAction:" & conNoError Then
                        'Error setting Schedule so exit
                        WriteScriptEngineLog "Error setting schedule, script aborted."
                    End If
                    Exit Do
                End If
            Case "SLEEP"
                'Wait for the appropriate time to continue
                Call SleepAction
                WriteScriptEngineLog "Sleep action finished."
            Case "END", "EXIT", "STOP"
                'No more actions in this script so exit
                Exit Do
            Case "LOG"
                'Read on to find a logfile to use and start logging
                If booLoggingFlag = True Then
                    'If logging is already on then ignore this action and read the next line instead
                    Line Input #lngScriptFileNumber, strCurrentScriptLine
                Else
                    Call LogAction
                End If
            Case "NOLOG"
                'Stop logging, then read the next line
                WriteScriptEngineLog "Log file closed."
                Close #lngLogFileNumber
                booLoggingFlag = False
                Line Input #lngScriptFileNumber, strCurrentScriptLine
            Case "LAUNCH"
                Call LaunchAction
            Case "FILE"
                Call FileAction
            Case "DIRECTORY"
                Call DirectoryAction
            Case "SCRIPT"
                Call ScriptAction
            Case "BEEP"
                'Make a noise
                Beep
                Beep
                Line Input #lngScriptFileNumber, strCurrentScriptLine
            'Case "NEW_ACTION"
            '    Call BoilerplateNewAction
            Case Else
                'This is the default action type, _
                    which reads for a server, source, destination and options
                Call StandardAction
            End Select
        Else
            'Read the next line if there was no action for this line
            Line Input #lngScriptFileNumber, strCurrentScriptLine
        End If
    Loop
    'Close the script file
    Close #lngScriptFileNumber
    WriteScriptEngineLog "Script file closed."
    strLastError = conSub & ":" & conNoError
    Beep
Exit Sub
'Error Handler
ProcessScript_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        WriteScriptEngineLog "Unexpected error : " & Err.Description & ", " & conClass & " : " & conSub
    End Select
    strLastError = conSub & ":" & conUndefinedError
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Exit Sub
End Sub

Private Sub BoilerplateNewAction()
'Created by Robin G Brown, 30/7/97
'Read the script file for an action
'Function Declares
Const conSub = "BoilerplateNewAction"
Dim intAssignPosition As Integer
Dim strArgument As String
Dim strParameter As String
Dim strActionValue As String
    'Error Trap
    On Error GoTo BoilerplateNewAction_ErrorHandler
    'Read the script file until we get to the next action line, or it ends
    Do While Not EOF(lngScriptFileNumber)
        'Read the next line
        Line Input #lngScriptFileNumber, strCurrentScriptLine
        If InStr(strCurrentScriptLine, "[") = 1 And InStr(strCurrentScriptLine, "]") > 0 Then
            'Action definition line so : _
                carry out the log action then exit this loop, _
                without reading the next line so that the outer routine will pick it up
            Exit Do
        ElseIf InStr(strCurrentScriptLine, "=") > 0 Then
            'Parameter definition line, process this line
            intAssignPosition = InStr(strCurrentScriptLine, "=")
            strArgument = Left$(strCurrentScriptLine, intAssignPosition - 1)
            strParameter = Right$(strCurrentScriptLine, Len(strCurrentScriptLine) - intAssignPosition)
            Select Case LCase$(strArgument)
            Case "ARGUMENT"
                'Set the name of the server to use
                strActionValue = Trim$(strParameter)
            Case Else
                'Not a valid script element, so ignore
            End Select
        Else
            'Something else - so ignore
        End If
    Loop
    'CODE_HERE_FOR strActionValue
    WriteScriptEngineLog "BoilerplateNewAction completed."
    strLastError = conSub & ":" & conNoError
Exit Sub
'Error Handler
BoilerplateNewAction_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        WriteScriptEngineLog "Unexpected error : " & Err.Description & ", " & conClass & " : " & conSub
    End Select
    strLastError = conSub & ":" & conUndefinedError
    Exit Sub
End Sub

Private Sub StandardAction()
'Created by Robin G Brown, 21/7/97
'Read the script file for a standard action, then execute it
'Function Declares
Const conSub = "StandardAction"
Dim intAssignPosition As Integer
Dim strArgument As String
Dim strParameter As String
Dim actStandard As EngineAction
    'Error Trap
    On Error GoTo StandardAction_ErrorHandler
    With actStandard
        .ServerName = ""
        .SourceName = ""
        .DestinationName = ""
        .OptionList = ""
    End With
    'Read the script file until we get to the next action line, or it ends
    Do While Not EOF(lngScriptFileNumber)
        'Read the next line
        Line Input #lngScriptFileNumber, strCurrentScriptLine
        If InStr(strCurrentScriptLine, "[") = 1 And InStr(strCurrentScriptLine, "]") > 0 Then
            'Action definition line so exit this loop, without reading the next line _
                so that the outer routine will pick it up
            Exit Do
        ElseIf InStr(strCurrentScriptLine, "=") > 0 Then
            'Parameter definition line, process this line
            intAssignPosition = InStr(strCurrentScriptLine, "=")
            strArgument = Left$(strCurrentScriptLine, intAssignPosition - 1)
            strParameter = Right$(strCurrentScriptLine, Len(strCurrentScriptLine) - intAssignPosition)
            Select Case LCase$(strArgument)
            Case "server"
                'Set the name of the server to use
                actStandard.ServerName = Trim$(strParameter)
            Case "source"
                'Set the name of the source datafile to use
                actStandard.SourceName = Trim$(strParameter)
            Case "destination"
                'Set the name of the destination file to use
                actStandard.DestinationName = Trim$(strParameter)
            Case "options"
                'Set the name of the destination file to use
                actStandard.OptionList = Trim$(strParameter)
            Case Else
                'Not a valid script element, so ignore
            End Select
        Else
            'Something else - so ignore
        End If
    Loop
    'Must have at least Server and Source data to execute an action
    If actStandard.ServerName = "" Or actStandard.SourceName = "" Then
        strLastError = conSub & ":" & conMissingParametersError
        Exit Sub
    Else
        Call StandardExecuteAction(actStandard)
        strLastError = conSub & ":" & conNoError
    End If
Exit Sub
'Error Handler
StandardAction_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        WriteScriptEngineLog "Unexpected error : " & Err.Description & ", " & conClass & " : " & conSub
    End Select
    strLastError = conSub & ":" & conUndefinedError
    Exit Sub
End Sub

Private Sub StandardExecuteAction(ByRef actExecute As EngineAction)
'Created by Robin G Brown, 11/7/97
'Run the three standard server functions
'Rather than use objServer for this action, we use objStandardServer _
    which is defined as having the IServerScriptAbs interface, _
    thus we can use the defined IServerScriptAbs methods _
    and avoid late binding, which should improve performance
'Function Declares
Const conSub = "StandardExecuteAction"
Dim objStandardServer As IServerScriptAbs
    'Error Trap
    On Error GoTo StandardExecuteAction_ErrorHandler
    'Create the appropriate server object
    WriteScriptEngineLog "Creating : " & actExecute.ServerName
    Set objStandardServer = Nothing
    Set objStandardServer = CreateObject(actExecute.ServerName)
    'Execute the three standard server methods
    WriteScriptEngineLog "Initialize : " & actExecute.ServerName
    If actExecute.OptionList <> "" Then
        lngReturn = objStandardServer.ServerInitialize(actExecute.SourceName, actExecute.DestinationName, actExecute.OptionList)
    Else
        lngReturn = objStandardServer.ServerInitialize(actExecute.SourceName, actExecute.DestinationName)
    End If
    If lngReturn <> 0 Then
        'Return the error number to the calling routine
        WriteScriptEngineLog "Initialize error = " & lngReturn
        Exit Sub
    End If
    WriteScriptEngineLog "Process : " & actExecute.ServerName
    lngReturn = objStandardServer.ServerProcess()
    If lngReturn <> 0 Then
        'Return the error number to the calling routine
        WriteScriptEngineLog "Process error = " & lngReturn
        Exit Sub
    End If
    WriteScriptEngineLog "Close : " & actExecute.ServerName
    lngReturn = objStandardServer.ServerClose()
    If lngReturn <> 0 Then
        'Return the error number to the calling routine
        WriteScriptEngineLog "Close error = " & lngReturn
        Exit Sub
    End If
    'Close the server
    Set objStandardServer = Nothing
    strLastError = conSub & ":" & conNoError
Exit Sub
'Error Handler
StandardExecuteAction_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        WriteScriptEngineLog "Unexpected error : " & Err.Description & ", " & conClass & " : " & conSub
    End Select
    strLastError = conSub & ":" & conUndefinedError
    Exit Sub
End Sub

Private Sub LaunchAction()
'Created by Robin G Brown, 23/7/97
'Read the script file for a launch action, then execute it
'Function Declares
Const conSub = "LaunchAction"
Dim intAssignPosition As Integer
Dim strArgument As String
Dim strParameter As String
Dim strCommandLine As String
    'Error Trap
    On Error GoTo LaunchAction_ErrorHandler
    'Read the script file until we get to the next action line, or it ends
    Do While Not EOF(lngScriptFileNumber)
        'Read the next line
        Line Input #lngScriptFileNumber, strCurrentScriptLine
        If InStr(strCurrentScriptLine, "[") = 1 And InStr(strCurrentScriptLine, "]") > 0 Then
            'Action definition line so exit this loop, without reading the next line _
                so that the outer routine will pick it up
            Exit Do
        ElseIf InStr(strCurrentScriptLine, "=") > 0 Then
            'Parameter definition line, process this line
            intAssignPosition = InStr(strCurrentScriptLine, "=")
            strArgument = Left$(strCurrentScriptLine, intAssignPosition - 1)
            strParameter = Right$(strCurrentScriptLine, Len(strCurrentScriptLine) - intAssignPosition)
            Select Case LCase$(strArgument)
            Case "program"
                'Set the name of the destination file to use
                strCommandLine = Trim$(strParameter)
            Case Else
                'Not a valid script element, so ignore
            End Select
        Else
            'Something else - so ignore
        End If
    Loop
    'Run the program
    If strCommandLine <> "" Then
        WriteScriptEngineLog "Launch program : " & strCommandLine
        Shell strCommandLine
    End If
    strLastError = conSub & ":" & conNoError
Exit Sub
'Error Handler
LaunchAction_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        WriteScriptEngineLog "Unexpected error : " & Err.Description & ", " & conClass & " : " & conSub
    End Select
    strLastError = conSub & ":" & conUndefinedError
    Exit Sub
End Sub

Private Sub SleepAction()
'Created by Robin G Brown, 21/7/97
'Read the script file for a date/time then wait until that date/time
'Function Declares
Const conSub = "SleepAction"
Dim intAssignPosition As Integer
Dim strArgument As String
Dim strParameter As String
Dim strTime As String
Dim lngWaitTime As Long
    'Error Trap
    On Error GoTo SleepAction_ErrorHandler
    'Read the script file until we get to the next action line, or it ends
    dteScheduledTime = 0
    Do While Not EOF(lngScriptFileNumber)
        'Read the next line
        Line Input #lngScriptFileNumber, strCurrentScriptLine
        If InStr(strCurrentScriptLine, "[") = 1 And InStr(strCurrentScriptLine, "]") > 0 Then
            'Action definition line so exit this loop, without reading the next line _
                so that the outer routine will pick it up
            Exit Do
        ElseIf InStr(strCurrentScriptLine, "=") > 0 Then
            'Parameter definition line, process this line
            intAssignPosition = InStr(strCurrentScriptLine, "=")
            strArgument = Left$(strCurrentScriptLine, intAssignPosition - 1)
            strParameter = Right$(strCurrentScriptLine, Len(strCurrentScriptLine) - intAssignPosition)
            Select Case LCase$(strArgument)
            Case "sleep"
                'Set the time of the schedule
                strTime = Trim$(strParameter)
                dteScheduledTime = TimeValue(strTime)
                If dteScheduledTime = 0 Then
                    'No time information so cannot continue
                    strLastError = conSub & ":" & conIncorrectScheduleError
                    Exit Sub
                End If
                'We have time information so there is no need to read further
                Exit Do
            Case Else
                'Not a valid script element, so ignore
            End Select
        Else
            'Something else - so ignore
        End If
    Loop
    'Wait until the right time
    If dteScheduledTime = 0 Then
        strLastError = conSub & ":" & conIncorrectScheduleError
        Exit Sub
    Else
        lngWaitTime = DateDiff("s", TimeValue(Now), dteScheduledTime)
        Set alaSleeper = CreateObject("AlarmTimer.Sleeper")
        Call alaSleeper.GoToSleep(lngWaitTime)
        Set alaSleeper = Nothing
        strLastError = conSub & ":" & conNoError
    End If
Exit Sub
'Error Handler
SleepAction_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        WriteScriptEngineLog "Unexpected error : " & Err.Description & ", " & conClass & " : " & conSub
    End Select
    strLastError = conSub & ":" & conUndefinedError
    Exit Sub
End Sub

Private Sub ScheduleAction()
'Created by Robin G Brown, 22/7/97
'Read the script file for a date/time then _
    set up a callback to activate the Script Engine _
    with that script at the appropriate date/time
'Function Declares
Const conSub = "ScheduleAction"
Dim intAssignPosition As Integer
Dim strArgument As String
Dim strParameter As String
Dim strTime As String
Dim lngWaitTime As Long
    'Error Trap
    On Error GoTo ScheduleAction_ErrorHandler
    'Read the script file until we get to the next action line, or it ends
    dteScheduledTime = 0
    Do While Not EOF(lngScriptFileNumber)
        'Read the next line
        Line Input #lngScriptFileNumber, strCurrentScriptLine
        If InStr(strCurrentScriptLine, "[") = 1 And InStr(strCurrentScriptLine, "]") > 0 Then
            'Action definition line so exit this loop, without reading the next line _
                so that the outer routine will pick it up
            Exit Do
        ElseIf InStr(strCurrentScriptLine, "=") > 0 Then
            'Parameter definition line, process this line
            intAssignPosition = InStr(strCurrentScriptLine, "=")
            strArgument = Left$(strCurrentScriptLine, intAssignPosition - 1)
            strParameter = Right$(strCurrentScriptLine, Len(strCurrentScriptLine) - intAssignPosition)
            Select Case LCase$(strArgument)
            Case "schedule"
                'Set the time of the schedule
                strTime = Trim$(strParameter)
                dteScheduledTime = CDate(strTime)
                If dteScheduledTime = 0 Then
                    'No date/time information so cannot continue
                    strLastError = conSub & ":" & conIncorrectScheduleError
                    Exit Sub
                End If
            Case Else
                'Not a valid script element, so ignore
            End Select
        Else
            'Something else - so ignore
        End If
    Loop
    'Wait until the right time
    If dteScheduledTime = 0 Then
        WriteScriptEngineLog "Schedule action set incorrectly, continuing."
        strLastError = conSub & ":" & conIncorrectScheduleError
        Exit Sub
    Else
        lngWaitTime = DateDiff("s", Now, dteScheduledTime)
        Set alaScheduler = New Alarm 'CreateObject("AlarmTimer.Alarm")
        Call alaScheduler.SetAlarm(Nothing, lngWaitTime)
        booAlarmWaiting = True
        'Must not set alaSchedule to nothing as this would cancel the AlarmRing event
        'Set alaScheduler = Nothing
        WriteScriptEngineLog "Alarm set for " & Format$(dteScheduledTime, "General Date")
        strLastError = conSub & ":" & conNoError
    End If
Exit Sub
'Error Handler
ScheduleAction_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        WriteScriptEngineLog "Unexpected error : " & Err.Description & ", " & conClass & " : " & conSub
    End Select
    strLastError = conSub & ":" & conUndefinedError
    Exit Sub
End Sub

Property Get AlarmWaiting() As Boolean
'Created by Robin G Brown, 22/7/97
'Check to see if there is an alarm waiting
'Function Declares
    'Error Trap
    On Error Resume Next
    AlarmWaiting = booAlarmWaiting
End Property

Public Sub CancelAlarm()
'Created by Robin G Brown, 22/7/97
'Cancel any waiting alarm, script will not be executed
'Sub Declares
    'Error Trap
    On Error Resume Next
    alaScheduler.CancelAlarm
    booAlarmWaiting = False
    Set alaScheduler = Nothing
    WriteScriptEngineLog "Alarm cancelled"
End Sub

Private Sub FileAction()
'Created by Robin G Brown, 30/7/97
'Read the script file for an action'Function Declares
Const conSub = "FileAction"
Dim intAssignPosition As Integer
Dim strArgument As String
Dim strParameter As String
Dim strSourceValue As String
Dim strDestinationValue As String
Dim strOptionsValue As String
    'Error Trap
    On Error GoTo FileAction_ErrorHandler
    'Read the script file until we get to the next action line, or it ends
    Do While Not EOF(lngScriptFileNumber)
        'Read the next line
        Line Input #lngScriptFileNumber, strCurrentScriptLine
        If InStr(strCurrentScriptLine, "[") = 1 And InStr(strCurrentScriptLine, "]") > 0 Then
            'Action definition line so : _
                carry out the action then exit this loop, _
                without reading the next line so that the outer routine will pick it up
            Exit Do
        ElseIf InStr(strCurrentScriptLine, "=") > 0 Then
            'Parameter definition line, process this line
            intAssignPosition = InStr(strCurrentScriptLine, "=")
            strArgument = Left$(strCurrentScriptLine, intAssignPosition - 1)
            strParameter = Right$(strCurrentScriptLine, Len(strCurrentScriptLine) - intAssignPosition)
            Select Case LCase$(strArgument)
            Case "options"
                'Set the options
                strOptionsValue = Trim$(strParameter)
            Case "source"
                'Set the source file
                strSourceValue = Trim$(strParameter)
            Case "destination"
                'Set the destination file
                strDestinationValue = Trim$(strParameter)
            Case Else
                'Not a valid script element, so ignore
            End Select
        Else
            'Something else - so ignore
        End If
    Loop
    'Create the appropriate server object
    WriteScriptEngineLog "Creating : FileServer.Server"
    Set objServer = Nothing
    Set objServer = CreateObject("FileServer.Server")
    'Execute the three standard server methods
    Select Case strOptionsValue
    Case "copy"
        Call objServer.CopyFile(strSourceValue, strDestinationValue)
    Case "move"
        Call objServer.MoveFile(strSourceValue, strDestinationValue)
    Case "delete"
        Call objServer.DeleteFile(strSourceValue)
    Case "rename"
        Call objServer.RenameFile(strSourceValue, strDestinationValue)
    Case Else
        'Error - so ignore
    End Select
    WriteScriptEngineLog "FileAction:" & strOptionsValue & " completed."
    Set objServer = Nothing
    strLastError = conSub & ":" & conNoError
Exit Sub
'Error Handler
FileAction_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        WriteScriptEngineLog "Unexpected error : " & Err.Description & ", " & conClass & " : " & conSub
    End Select
    strLastError = conSub & ":" & conUndefinedError
    Exit Sub
End Sub

Private Sub DirectoryAction()
'Created by Robin G Brown, 1/8/97
'Read the script file for an action
'Function Declares
Const conSub = "DirectoryAction"
Dim intAssignPosition As Integer
Dim strArgument As String
Dim strParameter As String
Dim strSourceValue As String
Dim strDestinationValue As String
Dim strOptionsValue As String
    'Error Trap
    On Error GoTo DirectoryAction_ErrorHandler
    'Read the script file until we get to the next action line, or it ends
    Do While Not EOF(lngScriptFileNumber)
        'Read the next line
        Line Input #lngScriptFileNumber, strCurrentScriptLine
        If InStr(strCurrentScriptLine, "[") = 1 And InStr(strCurrentScriptLine, "]") > 0 Then
            'Action definition line so : _
                carry out the action then exit this loop, _
                without reading the next line so that the outer routine will pick it up
            Exit Do
        ElseIf InStr(strCurrentScriptLine, "=") > 0 Then
            'Parameter definition line, process this line
            intAssignPosition = InStr(strCurrentScriptLine, "=")
            strArgument = Left$(strCurrentScriptLine, intAssignPosition - 1)
            strParameter = Right$(strCurrentScriptLine, Len(strCurrentScriptLine) - intAssignPosition)
            Select Case LCase$(strArgument)
            Case "options"
                'Set the options
                strOptionsValue = Trim$(strParameter)
            Case "source"
                'Set the source file
                strSourceValue = Trim$(strParameter)
            Case "destination"
                'Set the destination file
                strDestinationValue = Trim$(strParameter)
            Case Else
                'Not a valid script element, so ignore
            End Select
        Else
            'Something else - so ignore
        End If
    Loop
    'Create the appropriate server object
    WriteScriptEngineLog "Creating : FileServer.Server"
    Set objServer = Nothing
    Set objServer = CreateObject("FileServer.Server")
    'Execute the three standard server methods
    Select Case strOptionsValue
    Case "create"
        Call objServer.CreateDirectory(strSourceValue)
    Case "delete"
        Call objServer.DeleteDirectory(strSourceValue)
    Case "rename"
        Call objServer.RenameDirectory(strSourceValue, strDestinationValue)
    Case Else
        'Error - so ignore
    End Select
    WriteScriptEngineLog "DirectoryAction:" & strOptionsValue & " completed."
    Set objServer = Nothing
    strLastError = conSub & ":" & conNoError
Exit Sub
'Error Handler
DirectoryAction_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        WriteScriptEngineLog "Unexpected error : " & Err.Description & ", " & conClass & " : " & conSub
    End Select
    strLastError = conSub & ":" & conUndefinedError
    Exit Sub
End Sub

Private Sub ScriptAction()
'Created by Robin G Brown, 30/7/97
'Process another Script!
'Function Declares
Const conSub = "ScriptAction"
Dim intAssignPosition As Integer
Dim strArgument As String
Dim strParameter As String
Dim strScriptFile As String
Dim sceNewEngine As Engine
    'Error Trap
    On Error GoTo ScriptAction_ErrorHandler
    'Read the script file until we get to the next action line, or it ends
    Do While Not EOF(lngScriptFileNumber)
        'Read the next line
        Line Input #lngScriptFileNumber, strCurrentScriptLine
        If InStr(strCurrentScriptLine, "[") = 1 And InStr(strCurrentScriptLine, "]") > 0 Then
            'Action definition line so : _
                carry out the log action then exit this loop, _
                without reading the next line so that the outer routine will pick it up
            Exit Do
        ElseIf InStr(strCurrentScriptLine, "=") > 0 Then
            'Parameter definition line, process this line
            intAssignPosition = InStr(strCurrentScriptLine, "=")
            strArgument = Left$(strCurrentScriptLine, intAssignPosition - 1)
            strParameter = Right$(strCurrentScriptLine, Len(strCurrentScriptLine) - intAssignPosition)
            Select Case LCase$(strArgument)
            Case "script"
                'Set the name of the server to use
                strScriptFile = Trim$(strParameter)
            Case Else
                'Not a valid script element, so ignore
            End Select
        Else
            'Something else - so ignore
        End If
    Loop
    'Create another Engine object and use it to run the script
    Set sceNewEngine = New Engine
    Call sceNewEngine.RunFile(strScriptFile)
    Set sceNewEngine = Nothing
    WriteScriptEngineLog "ScriptAction for file : " & strScriptFile & " completed."
    strLastError = conSub & ":" & conNoError
Exit Sub
'Error Handler
ScriptAction_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        WriteScriptEngineLog "Unexpected error : " & Err.Description & ", " & conClass & " : " & conSub
    End Select
    strLastError = conSub & ":" & conUndefinedError
    Exit Sub
End Sub

Private Sub LogAction()
'Created by Robin G Brown, 21/7/97
'Read the script file for a log action
'Function Declares
Const conSub = "LogAction"
Dim intAssignPosition As Integer
Dim strArgument As String
Dim strParameter As String
    'Error Trap
    On Error GoTo LogAction_ErrorHandler
    'Read the script file until we get to the next action line, or it ends
    Do While Not EOF(lngScriptFileNumber)
        'Read the next line
        Line Input #lngScriptFileNumber, strCurrentScriptLine
        If InStr(strCurrentScriptLine, "[") = 1 And InStr(strCurrentScriptLine, "]") > 0 Then
            'Action definition line so : _
                carry out the log action then exit this loop, _
                without reading the next line so that the outer routine will pick it up
            Exit Do
        ElseIf InStr(strCurrentScriptLine, "=") > 0 Then
            'Parameter definition line, process this line
            intAssignPosition = InStr(strCurrentScriptLine, "=")
            strArgument = Left$(strCurrentScriptLine, intAssignPosition - 1)
            strParameter = Right$(strCurrentScriptLine, Len(strCurrentScriptLine) - intAssignPosition)
            Select Case LCase$(strArgument)
            Case "logfile"
                'Set the name of the server to use
                strLogFilename = Trim$(strParameter)
            Case Else
                'Not a valid script element, so ignore
            End Select
        Else
            'Something else - so ignore
        End If
    Loop
    booLoggingFlag = True
    lngLogFileNumber = FreeFile()
    Open strLogFilename For Output As #lngLogFileNumber
    Write #lngLogFileNumber, Now, "Log file opened."
    Close #lngLogFileNumber
    strLastError = conSub & ":" & conNoError
Exit Sub
'Error Handler
LogAction_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        WriteScriptEngineLog "Unexpected error : " & Err.Description & ", " & conClass & " : " & conSub
    End Select
    strLastError = conSub & ":" & conUndefinedError
    Exit Sub
End Sub

Private Sub WriteScriptEngineLog(ByVal strLogItem As String)
'Created by Robin G Brown, 21/7/97
'Write a line to the logfile
'Sub Declares
    'Error Trap
    On Error Resume Next
    'Check to see if logging is required
    If booLoggingFlag = True Then
        'Write a line to the file with the time
        lngLogFileNumber = FreeFile()
        Open strLogFilename For Append As #lngLogFileNumber
        Write #lngLogFileNumber, Now, strLogItem
        Write #lngLogFileNumber, Now, "Last Error = "; LastErrorMessage
        Close #lngLogFileNumber
    End If
End Sub

Property Get LastErrorMessage() As String
'Created by Robin G Brown, 6/8/97
'Return the latest internal error
'Function Declares
    'Error Trap
    On Error Resume Next
    LastErrorMessage = strLastError
End Property

Private Sub Class_Terminate()
'Created by Robin G Brown, 21/7/97
'Default behaviour
'Sub Declares
    'Error Trap
    On Error Resume Next
    'Close files, clear objects, etc.
    Close #lngScriptFileNumber
    If Err = 0 Then
        WriteScriptEngineLog "Script file closed."
    End If
    WriteScriptEngineLog "Script Engine terminated."
    Set objServer = Nothing
    Set alaSleeper = Nothing
    Set alaScheduler = Nothing
End Sub
