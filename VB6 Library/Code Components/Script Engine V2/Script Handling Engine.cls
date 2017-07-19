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
'   Script File Format Information is in the AnoibjectModel 'module'
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---Events
'Instances of the engine object will need to be declared WithEvents
'Script update is used to inform the calling appliction _
    that part of a script has been completed
Public Event ScriptUpdate(ByVal strMessage As String, ByVal intPercentDone As Integer)
'ScriptCompleted is used to inform the calling application _
    that the requested script has finished execution
Public Event ScriptCompleted()
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
Private WithEvents alaScheduler As Alarm
Attribute alaScheduler.VB_VarHelpID = -1
Private scfScript As script
Private scfSubScript As script
'Private properties
Private lngReturn As Long
Private strLogFilename As String
Private strLastError As String
Private lngLogFileNumber As Long
Private dteScheduledTime As Date
Private booLoggingFlag As Boolean
Private booAlarmWaiting As Boolean
Private booScriptRunning As Boolean
'Internal Error Codes
Private Const conNoError = "No Error"
Private Const conUndefinedError = "Undefined"
Private Const conLoggingActiveError = "Logging already active"
Private Const conIncorrectScheduleError = "Incorrect Schedule"
Private Const conMissingParametersError = "Missing Parameters"
Private Const conStdServInitFailError = "Standard Server initialization failed"
Private Const conStdServProcFailError = "Standard Server processing failed"
Private Const conStdServCloseFailError = "Standard Server close failed"

Private Sub Class_Initialize()
'Created by Robin G Brown, 21/7/97
'Default behaviour
'Sub Declares
    'Error Trap
    On Error Resume Next
    booAlarmWaiting = False
    strLastError = conNoError
    'Initialize is only called once, so scfscript must be new each time a script is run
    'Set scfScript = New Script 'this line is now in the RunFile/RunText procedures
    If Err.Number <> 0 Then
        Err.Raise Err.Number, conClass & ":Initialize", "Unexpected error : " & Err.Description
    End If
End Sub

'-----------------------------------------------------------------------------
'   Event Handlers
'-----------------------------------------------------------------------------

Private Sub alaScheduler_AlarmRing()
'Created by Robin G Brown, 22/7/97
'This is the return event from the Alarm object
'Sub Declares
    'Error Trap
    On Error Resume Next
    'Open the script and run it
    WriteScriptEngineLog "Alarm rung."
    Call ProcessScript
    'Set booAlarmWaiting after the script has run so that it ignores the Schedule action
    booAlarmWaiting = False
    Set alaScheduler = Nothing
'    If Err.Number <> 0 Then
'        'There was an error, probably when running the script, _
'            so put up a message for the user informing them of the problem
'        MsgBox "There was an error (Error: " & Err.Number & ", '" & Err.Description & "') while running a scheduled script, please retry.", vbOKOnly + vbExclamation, "Script Engine"
'    End If
End Sub

'-----------------------------------------------------------------------------
'   Public Functions, Methods and Properties
'-----------------------------------------------------------------------------

Property Get AlarmWaiting() As Boolean
'Created by Robin G Brown, 22/7/97
'Check to see if there is an alarm waiting
'Function Declares
    'Error Trap
    On Error Resume Next
    AlarmWaiting = booAlarmWaiting
End Property

Property Get ScriptRunning() As Boolean
'Created by Robin G Brown, 22/7/97
'Check to see if there is a script running
'Function Declares
    'Error Trap
    On Error Resume Next
    ScriptRunning = booScriptRunning
End Property

Property Get LastErrorMessage() As String
'Created by Robin G Brown, 6/8/97
'Return the latest internal error
'Function Declares
    'Error Trap
    On Error Resume Next
    LastErrorMessage = strLastError
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

Public Sub RunText(ByVal strScriptText As String)
'Created by Robin G Brown, 21/7/97
'Set the script text and set the script engine to process it, _
    this sub returns straightaway and sets a timer to start processing in 3 seconds
'Sub Declares
Const conSub = "RunText"
    'Error Trap
    On Error Resume Next
    WriteScriptEngineLog "RunText method initiated"
    If booScriptRunning = True Then
        'Prevent re-entrancy!
        WriteScriptEngineLog "Attempted re-entry prevented."
        Err.Raise vbObjectError, conClass & ":" & conSub, "Attempting to run another script, while a script is waiting to be run, cannot continue."
        Exit Sub
    End If
    'Create a new script object each time! - RGB/11/9/97
    Set scfScript = Nothing
    Set scfScript = New script
    'Error Trap changed to immediate error handling for rest of procedure
    On Error GoTo RunText_ErrorHandler
    'Read the text into the Script object
    scfScript.Text = strScriptText
    'Create and set an alarm which will process the script
    Set alaScheduler = New Alarm 'CreateObject("AlarmTimer.Alarm")
    Call alaScheduler.SetAlarm(Nothing, 3)
    booScriptRunning = True
    booAlarmWaiting = True
    'Must not set alaSchedule to nothing as this would cancel the AlarmRing event
    'Set alaScheduler = Nothing
    strLastError = conSub & ":" & conNoError
Exit Sub
RunText_ErrorHandler:
    WriteScriptEngineLog "Unexpected error : " & Err.Description & ", " & conClass & " : " & conSub
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub RunFile(ByVal strScriptFileName As String)
'Created by Robin G Brown, 21/7/97
'Set the script file and set the script engine to process it, _
    this sub returns straightaway and sets a timer to start processing in 3 seconds
'Sub Declares
Const conSub = "RunFile"
    'Error Trap
    On Error Resume Next
    WriteScriptEngineLog "RunFile method initiated"
    If booScriptRunning = True Then
        'Prevent re-entrancy!
        WriteScriptEngineLog "Attempted re-entry prevented."
        Err.Raise vbObjectError, conClass & ":" & conSub, "Attempting to run another script, while a script is waiting to be run, cannot continue."
        Exit Sub
    End If
    'Create a new script object each time! - RGB/11/9/97
    Set scfScript = Nothing
    Set scfScript = New script
    'Error Trap changed to immediate error handling for rest of procedure
    On Error GoTo RunFile_ErrorHandler
    'Read the file into the Script object
    Call scfScript.FileOpen(strScriptFileName)
    'Create and set an alarm which will process the script
    Set alaScheduler = New Alarm 'CreateObject("AlarmTimer.Alarm")
    Call alaScheduler.SetAlarm(Nothing, 3)
    booScriptRunning = True
    booAlarmWaiting = True
    'Must not set alaSchedule to nothing as this would cancel the AlarmRing event
    'Set alaScheduler = Nothing
    strLastError = conSub & ":" & conNoError
Exit Sub
RunFile_ErrorHandler:
    WriteScriptEngineLog "Unexpected error : " & Err.Description & ", " & conClass & " : " & conSub
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

'-----------------------------------------------------------------------------
'   Private Functions and Methods
'-----------------------------------------------------------------------------

Private Sub ProcessScript()
'Created by Robin G Brown, 10/7/97
'Read the 'Actions' in the Script object, set the variables and run the server
'Sub Declares
Const conSub = "ProcessScript"
Dim strActionType As String
Dim lngActionCounter As Long
Dim secAction As Section
    'Error Trap
    On Error GoTo ProcessScript_ErrorHandler
    'Now process actions until we run out or are told to stop
    For lngActionCounter = 1 To scfScript.ActionCount
        'Process this action
        Set secAction = scfScript.Action(lngActionCounter)
        Select Case secAction.Name
        Case "TEST"
            'MsgBox "Test of ScriptEngineV2.Engine(" & conSub & ") succesfull.", vbOKOnly, App.ProductName
            WriteScriptEngineLog "Test of ScriptEngineV2.Engine(" & conSub & ") succesfull."
        Case "SCHEDULE"
            RaiseEvent ScriptUpdate("Schedule set, waiting", 0)
            If booAlarmWaiting = True Then
                'If there is an alarm waiting then ignore the Schedule action
                 WriteScriptEngineLog "Alarm is set, ignoring Schedule action"
            Else
                'Exit this function setting up an alarm to continue later!
                Call ScheduleAction(secAction)
                If strLastError <> "ScheduleAction:" & conNoError Then
                    'Error setting Schedule so exit
                    WriteScriptEngineLog "Error setting schedule, script aborted."
                End If
                'Finish off and then exit
                WriteScriptEngineLog "Script scheduled for later processing, engine waiting."
                strLastError = conSub & ":" & conNoError
                'Exit here so that booScriptRunning is not set to false!
                booScriptRunning = True
                Exit Sub
            End If
        Case "SLEEP"
            RaiseEvent ScriptUpdate("Waiting", 0)
            'Wait for the appropriate time to continue
            Call SleepAction(secAction)
            WriteScriptEngineLog "Sleep action finished."
        Case "END", "EXIT", "STOP"
            RaiseEvent ScriptUpdate("End of Script", 100)
            'No more actions in this script so exit
            Exit For
        Case "LOG"
            RaiseEvent ScriptUpdate("Log started", 0)
            'Read on to find a logfile to use and start logging
            If booLoggingFlag = True Then
                'If logging is already on then ignore this action
            Else
                Call LogAction(secAction)
            End If
        Case "NOLOG"
            RaiseEvent ScriptUpdate("Log finished", 0)
            'Stop logging, then read the next line
            WriteScriptEngineLog "Log file closed."
            Close #lngLogFileNumber
            booLoggingFlag = False
        Case "LAUNCH"
            RaiseEvent ScriptUpdate("Program launched", 0)
            Call LaunchAction(secAction)
        Case "FILE"
            RaiseEvent ScriptUpdate("File action", 0)
            Call FileAction(secAction)
        Case "DIRECTORY"
            RaiseEvent ScriptUpdate("Directory action", 0)
            Call DirectoryAction(secAction)
        Case "SCRIPT"
            RaiseEvent ScriptUpdate("Sub script opened", 0)
            Call ScriptAction(secAction)
            RaiseEvent ScriptUpdate("Sub script finished", 0)
        Case "BEEP"
            RaiseEvent ScriptUpdate("Beep", 0)
            'Make a noise
            Beep
            Beep
        'Case "NEW_ACTION"
        '    RaiseEvent ScriptUpdate("NEW_ACTION", 0)
        '    Call BoilerplateNewAction
        Case Else
            'This is the default action type, _
                which reads for a server, source, destination and options
            RaiseEvent ScriptUpdate("Standard action : " & secAction.Name, 0)
            Call StandardAction(secAction)
        End Select
    Next
    WriteScriptEngineLog "Script completed."
    strLastError = conSub & ":" & conNoError
    booScriptRunning = False
    RaiseEvent ScriptCompleted
Exit Sub
'Error Handler
ProcessScript_ErrorHandler:
    WriteScriptEngineLog "Unexpected error : " & Err.Description & ", " & conClass & " : " & conSub
    strLastError = conSub & ":" & conUndefinedError
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Exit Sub
End Sub

Private Sub ProcessSubScript(ByRef scfOtherScript As script)
'Created by Robin G Brown, 10/7/97
'Read the 'Actions' in the scfOtherScript object, set the variables and run the servers
'Sub Declares
Const conSub = "ProcessSubScript"
Dim strActionType As String
Dim lngActionCounter As Long
Dim secAction As Section
    'Error Trap
    On Error GoTo ProcessSubScript_ErrorHandler
    'Now process actions until we run out or are told to stop
    For lngActionCounter = 1 To scfOtherScript.ActionCount
        'Process this action
        Set secAction = scfOtherScript.Action(lngActionCounter)
        Select Case secAction.Name
        Case "TEST"
            'MsgBox "Test of ScriptEngineV2.Engine(" & conSub & ") succesfull.", vbOKOnly, App.ProductName
            WriteScriptEngineLog "Test of ScriptEngineV2.Engine(" & conSub & ") succesfull."
        Case "SCHEDULE", "SLEEP", "SCRIPT"
            RaiseEvent ScriptUpdate(secAction.Name & " Action cancelled", 0)
            'Do nothing, these actions in subscripts must be ignored!
        Case "END", "EXIT", "STOP"
            RaiseEvent ScriptUpdate("End of Script", 100)
            'No more actions in this script so exit
            Exit For
        Case "LOG"
            RaiseEvent ScriptUpdate("Log started", 0)
            'Read on to find a logfile to use and start logging
            If booLoggingFlag = True Then
                'If logging is already on then ignore this action
            Else
                Call LogAction(secAction)
            End If
        Case "NOLOG"
            RaiseEvent ScriptUpdate("Log finished", 0)
            'Stop logging, then read the next line
            WriteScriptEngineLog "Log file closed."
            Close #lngLogFileNumber
            booLoggingFlag = False
        Case "LAUNCH"
            RaiseEvent ScriptUpdate("Program launched", 0)
            Call LaunchAction(secAction)
        Case "FILE"
            RaiseEvent ScriptUpdate("File action", 0)
            Call FileAction(secAction)
        Case "DIRECTORY"
            RaiseEvent ScriptUpdate("Directory action", 0)
            Call DirectoryAction(secAction)
        Case "BEEP"
            RaiseEvent ScriptUpdate("Beep", 0)
            'Make a noise
            Beep
            Beep
        'Case "NEW_ACTION"
        '    RaiseEvent ScriptUpdate("NEW_ACTION", 0)
        '    Call BoilerplateNewAction
        Case Else
            'This is the default action type, _
                which reads for a server, source, destination and options
            RaiseEvent ScriptUpdate("Standard action : " & secAction.Name, 0)
            Call StandardAction(secAction)
        End Select
    Next
    WriteScriptEngineLog "Subscript completed."
    strLastError = conSub & ":" & conNoError
    booScriptRunning = False
Exit Sub
'Error Handler
ProcessSubScript_ErrorHandler:
    WriteScriptEngineLog "Unexpected error : " & Err.Description & ", " & conClass & " : " & conSub
    strLastError = conSub & ":" & conUndefinedError
    Exit Sub
End Sub

Private Sub BoilerplateNewAction(ByRef secAction As Section)
'Created by Robin G Brown, 30/7/97
'Read the script file for an action
'Function Declares
Const conSub = "BoilerplateNewAction"
Dim strActionValue As String
    'Error Trap
    On Error GoTo BoilerplateNewAction_ErrorHandler
    'Read all the keys in the script object, then process the action
    With secAction
        strActionValue = Trim$(.KeyValue("ARGUMENT_HERE"))
        'REPEAT_AS_NECESSARY
        strActionValue = Trim$(.KeyValue("ARGUMENT_HERE"))
        'REPEAT_AS_NECESSARY
    End With
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

Private Sub StandardAction(ByRef secAction As Section)
'Created by Robin G Brown, 21/7/97
'Read the script file for a standard action, then execute it
'Function Declares
Const conSub = "StandardAction"
Dim lngKeyCounter As Long
Dim actStandard As EngineAction
Dim strValue As String
'Rather than use an undefined object for this action, we use objStandardServer _
    which is defined as having the IServerScriptAbs interface, _
    thus we can use the defined IServerScriptAbs methods _
    and avoid late binding, which should improve performance
Dim objStandardServer As IServerScriptAbs
    'Error Trap
    On Error GoTo StandardAction_ErrorHandler
    'Blank actStandard first
    With actStandard
        .ServerName = ""
        .SourceName = ""
        .DestinationName = ""
        .OptionList = ""
    End With
    'Read all the keys in the script object, then process the action
    With secAction
        actStandard.ServerName = Trim$(.KeyValue("server"))
        actStandard.SourceName = Trim$(.KeyValue("source"))
        actStandard.DestinationName = Trim$(.KeyValue("destination"))
        actStandard.OptionList = Trim$(.KeyValue("options"))
    End With
    With actStandard
        'Must have at least Server and Source data to execute an action
        If .ServerName = "" Or .SourceName = "" Then
            WriteScriptEngineLog "Missing parameters for Standard action, action failed."
            strLastError = conSub & ":" & conMissingParametersError
            Exit Sub
        Else
            'Create the appropriate server object
            WriteScriptEngineLog "Creating : " & .ServerName
            Set objStandardServer = Nothing
            Set objStandardServer = CreateObject(.ServerName)
            'Execute the three standard server methods
            WriteScriptEngineLog "Initialize : " & .ServerName
            If actStandard.OptionList <> "" Then
                lngReturn = objStandardServer.ServerInitialize(.SourceName, .DestinationName, .OptionList)
            Else
                lngReturn = objStandardServer.ServerInitialize(.SourceName, .DestinationName)
            End If
            If lngReturn <> 0 Then
                'Return the error number to the calling routine
                WriteScriptEngineLog "Initialize error = " & lngReturn
                strLastError = conSub & ":" & conStdServInitFailError
                Set objStandardServer = Nothing
                Exit Sub
            End If
            WriteScriptEngineLog "Process : " & .ServerName
            lngReturn = objStandardServer.ServerProcess()
            If lngReturn <> 0 Then
                'Return the error number to the calling routine
                WriteScriptEngineLog "Process error = " & lngReturn
                strLastError = conSub & ":" & conStdServProcFailError
                Set objStandardServer = Nothing
                Exit Sub
            End If
            WriteScriptEngineLog "Close : " & .ServerName
            lngReturn = objStandardServer.ServerClose()
            If lngReturn <> 0 Then
                'Return the error number to the calling routine
                WriteScriptEngineLog "Close error = " & lngReturn
                strLastError = conSub & ":" & conStdServCloseFailError
                Set objStandardServer = Nothing
                Exit Sub
            End If
            'Close the server
            Set objStandardServer = Nothing
            strLastError = conSub & ":" & conNoError
        End If
    End With
Exit Sub
'Error Handler
StandardAction_ErrorHandler:
    WriteScriptEngineLog "Unexpected error : " & Err.Description & ", " & conClass & " : " & conSub
    Set objStandardServer = Nothing
    strLastError = conSub & ":" & conUndefinedError
    Exit Sub
End Sub

Private Sub LaunchAction(ByRef secAction As Section)
'Created by Robin G Brown, 23/7/97
'Read the script file for a launch action, then execute it
'Function Declares
Const conSub = "LaunchAction"
Dim lngKeyCounter As Long
Dim strArgument As String
Dim strParameter As String
Dim strCommandLine As String
    'Error Trap
    On Error GoTo LaunchAction_ErrorHandler
    'Read all the keys in the script object, then process the action
    strCommandLine = Trim$(secAction.KeyValue("program"))
    'Run the program
    If strCommandLine <> "" Then
        WriteScriptEngineLog "Launch program with command line : " & strCommandLine
        Shell strCommandLine
    End If
    strLastError = conSub & ":" & conNoError
Exit Sub
'Error Handler
LaunchAction_ErrorHandler:
    WriteScriptEngineLog "Unexpected error : " & Err.Description & ", " & conClass & " : " & conSub
    strLastError = conSub & ":" & conUndefinedError
    Exit Sub
End Sub

Private Sub SleepAction(ByRef secAction As Section)
'Created by Robin G Brown, 21/7/97
'Read the script file for a date/time then wait until that date/time
'Function Declares
Const conSub = "SleepAction"
Dim alaSleeper As Sleeper
Dim strTime As String
Dim lngWaitTime As Long
    'Error Trap
    On Error GoTo SleepAction_ErrorHandler
    'Read all the keys in the script object, then process the action
    dteScheduledTime = 0
    'Set the time of the schedule
    strTime = Trim$(secAction.KeyValue("sleep"))
    dteScheduledTime = TimeValue(strTime)
    If dteScheduledTime = 0 Then
        'No time information so cannot continue
        WriteScriptEngineLog "Sleep action set incorrectly, continuing."
        strLastError = conSub & ":" & conIncorrectScheduleError
        Exit Sub
    End If
    'Wait until the right time
    If dteScheduledTime = 0 Then
        WriteScriptEngineLog "Sleep action set incorrectly, continuing."
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
    WriteScriptEngineLog "Unexpected error : " & Err.Description & ", " & conClass & " : " & conSub
    strLastError = conSub & ":" & conUndefinedError
    Exit Sub
End Sub

Private Sub ScheduleAction(ByRef secAction As Section)
'Created by Robin G Brown, 22/7/97
'Read the script file for a date/time then _
    set up a callback to activate the Script Engine _
    with that script at the appropriate date/time
'Function Declares
Const conSub = "ScheduleAction"
Dim strTime As String
Dim lngWaitTime As Long
    'Error Trap
    On Error GoTo ScheduleAction_ErrorHandler
    'Read the key in the section object, then process the action
    dteScheduledTime = 0
    'Set the time of the schedule
    strTime = Trim$(secAction.KeyValue("schedule"))
    dteScheduledTime = CDate(strTime)
    If dteScheduledTime = 0 Then
        'No date/time information so cannot continue
        strLastError = conSub & ":" & conIncorrectScheduleError
        Exit Sub
    End If
    'Wait until the right time
    If dteScheduledTime = 0 Then
        WriteScriptEngineLog "Schedule action set incorrectly, continuing."
        strLastError = conSub & ":" & conIncorrectScheduleError
        booAlarmWaiting = False
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
    WriteScriptEngineLog "Unexpected error : " & Err.Description & ", " & conClass & " : " & conSub
    strLastError = conSub & ":" & conUndefinedError
    Exit Sub
End Sub

Private Sub FileAction(ByRef secAction As Section)
'Created by Robin G Brown, 30/7/97
'Read the script file for an action
'Function Declares
Const conSub = "FileAction"
Dim strSourceValue As String
Dim strDestinationValue As String
Dim strOptionsValue As String
Dim objFileServer As FileServer.Server
    'Error Trap
    On Error GoTo FileAction_ErrorHandler
    'Read all the keys in the script object, then process the action
    With secAction
        strOptionsValue = Trim$(.KeyValue("options"))
        strSourceValue = Trim$(.KeyValue("source"))
        strDestinationValue = Trim$(.KeyValue("destination"))
    End With
    'Create the appropriate server object
    WriteScriptEngineLog "Creating : FileServer.Server"
    Set objFileServer = Nothing
    Set objFileServer = New FileServer.Server
    'Execute the three standard server methods
    Select Case strOptionsValue
    Case "copy"
        Call objFileServer.CopyFile(strSourceValue, strDestinationValue)
    Case "move"
        Call objFileServer.MoveFile(strSourceValue, strDestinationValue)
    Case "delete"
        Call objFileServer.DeleteFile(strSourceValue)
    Case "rename"
        Call objFileServer.RenameFile(strSourceValue, strDestinationValue)
    Case Else
        'Error - so ignore
    End Select
    WriteScriptEngineLog "FileAction:" & strOptionsValue & " completed."
    Set objFileServer = Nothing
    strLastError = conSub & ":" & conNoError
Exit Sub
'Error Handler
FileAction_ErrorHandler:
    WriteScriptEngineLog "Unexpected error : " & Err.Description & ", " & conClass & " : " & conSub
    Set objFileServer = Nothing
    strLastError = conSub & ":" & conUndefinedError
    Exit Sub
End Sub

Private Sub DirectoryAction(ByRef secAction As Section)
'Created by Robin G Brown, 1/8/97
'Read the script file for an action
'Function Declares
Const conSub = "DirectoryAction"
Dim strSourceValue As String
Dim strDestinationValue As String
Dim strOptionsValue As String
Dim objFileServer As FileServer.Server
    'Error Trap
    On Error GoTo DirectoryAction_ErrorHandler
    'Read all the keys in the script object, then process the action
    With secAction
        strSourceValue = Trim$(.KeyValue("source"))
        strDestinationValue = Trim$(.KeyValue("destination"))
        strOptionsValue = Trim$(.KeyValue("options"))
    End With
    'Create the appropriate server object
    WriteScriptEngineLog "Creating : FileServer.Server"
    Set objFileServer = New FileServer.Server
    'Execute the three standard server methods
    Select Case strOptionsValue
    Case "create"
        Call objFileServer.CreateDirectory(strSourceValue)
    Case "delete"
        Call objFileServer.DeleteDirectory(strSourceValue)
    Case "rename"
        Call objFileServer.RenameDirectory(strSourceValue, strDestinationValue)
    Case Else
        'Error - so ignore
    End Select
    WriteScriptEngineLog "DirectoryAction:" & strOptionsValue & " completed."
    Set objFileServer = Nothing
    strLastError = conSub & ":" & conNoError
Exit Sub
'Error Handler
DirectoryAction_ErrorHandler:
    WriteScriptEngineLog "Unexpected error : " & Err.Description & ", " & conClass & " : " & conSub
    Set objFileServer = Nothing
    strLastError = conSub & ":" & conUndefinedError
    Exit Sub
End Sub

Private Sub ScriptAction(ByRef secAction As Section)
'Created by Robin G Brown, 30/7/97
'Process another Script!
'Function Declares
Const conSub = "ScriptAction"
Dim strScriptFile As String
Dim scfSubScript As New script
    'Error Trap
    On Error GoTo ScriptAction_ErrorHandler
    'Read all the keys in the script object, then process the action
    strScriptFile = Trim$(secAction.KeyValue("script"))
    scfSubScript.FileOpen strScriptFile
    'Pass the new (sub)script object to ProcessSubScript
    Call ProcessSubScript(scfSubScript)
    Set scfSubScript = Nothing
    WriteScriptEngineLog "ScriptAction for file : " & strScriptFile & " completed."
    strLastError = conSub & ":" & conNoError
Exit Sub
'Error Handler
ScriptAction_ErrorHandler:
    WriteScriptEngineLog "Unexpected error : " & Err.Description & ", " & conClass & " : " & conSub
    strLastError = conSub & ":" & conUndefinedError
    Exit Sub
End Sub

Private Sub LogAction(ByRef secAction As Section)
'Created by Robin G Brown, 21/7/97
'Read the script file for a log action
'Function Declares
Const conSub = "LogAction"
Dim lngKeyCounter As Long
Dim strArgument As String
Dim strParameter As String
    'Error Trap
    On Error GoTo LogAction_ErrorHandler
    'Read all the keys in the script object, then process the action
    strLogFilename = Trim$(secAction.KeyValue("logfile"))
    booLoggingFlag = True
    lngLogFileNumber = FreeFile()
    Open strLogFilename For Output As #lngLogFileNumber
    Write #lngLogFileNumber, Format$(Now, "dd/mm/yy hh:nn:ss"), "Log file opened."
    Close #lngLogFileNumber
    strLastError = conSub & ":" & conNoError
Exit Sub
'Error Handler
LogAction_ErrorHandler:
    WriteScriptEngineLog "Unexpected error : " & Err.Description & ", " & conClass & " : " & conSub
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
        'Write strLogItem to the file with the time
        lngLogFileNumber = FreeFile()
        Open strLogFilename For Append As #lngLogFileNumber
        Write #lngLogFileNumber, Format$(Now, "dd/mm/yy hh:nn:ss"), strLogItem
        Close #lngLogFileNumber
    End If
End Sub

Private Sub Class_Terminate()
'Created by Robin G Brown, 21/7/97
'Default behaviour
'Sub Declares
    'Error Trap
    On Error Resume Next
    'Close files, clear objects, etc.
    WriteScriptEngineLog "Script Engine terminated."
    Set alaScheduler = Nothing
    Set scfScript = Nothing
End Sub
