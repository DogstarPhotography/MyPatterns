VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Builder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Created by Robin G Brown, 23rd July 1997
'-----------------------------------------------------------------------------
'   This class module contains code for _
    a script building object
'   Use bld as a prefix for this class
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "Builder"
Private Const conActionsQuery = "Actions"
Private Const conAllActionsQuery = "ActionsParametersValues"
Private Const conCommandsQuery = "Commands"
Private Const conServersQuery = "Servers"
Private Const conActionParameterQuery = "ActionParameter"
Private Const conParametersQuery = "Parameters"
Private Const conDefaultScriptDBName = "e:\janson green\data\scripts.mdb"
'Public properties
'Private properties
Private lngCounter As Long
Private lngReturn As Long
Private strActions() As String
'Database variables
Private wsScriptsWorkspace As Workspace
Private dbScriptsDatabase As Database
Private rsActions As Recordset
Private rsActionsQuery As Recordset
Private rsCommands As Recordset
Private rsParameters As Recordset
Private rsServers As Recordset
Private strScriptsDatabaseName As String

Private Sub Class_Initialize()
'Created by Robin G Brown, 28/7/97
'Open the database
Dim sdlDialog As Dialog
    'Error Trap
    On Error Resume Next
    Set sdlDialog = New Dialog
    strScriptsDatabaseName = GetSetting(App.Title, "Database", "Name", conDefaultScriptDBName)
    Call ConnectScriptsDatabase
    Do While Err.Number <> 0
        'First try to find DB
        lngReturn = MsgBox("Could not connect to database, this server may not run correctly, click OK to select a database, Cancel to exit.", vbOKCancel, "Builder Database Connect")
        If lngReturn = vbOK Then
            sdlDialog.DatabaseName = strScriptsDatabaseName
            sdlDialog.ShowSelectDatabase
            strScriptsDatabaseName = sdlDialog.DatabaseName
        Else
            Err.Raise vbObjectError, conModule, "Initialization of Script Engine Builder failed due to database error number : " & lngReturn
        End If
        Err.Clear
        Call ConnectScriptsDatabase
    Loop
    Set rsActionsQuery = dbScriptsDatabase.OpenRecordset(conAllActionsQuery, dbOpenDynaset)
    Set rsActions = dbScriptsDatabase.OpenRecordset(conActionsQuery, dbOpenDynaset)
    Set sdlDialog = Nothing
End Sub

Private Sub ConnectScriptsDatabase()
'Created by Robin G Brown, 28/7/97
'This function is designed to connect to the Scripts Access database
'Returns true if succesful, false if an error
    'gstrDBName must be a full path and file name
    'Error trap
    On Error GoTo ConnectScriptsDatabase_ErrorHandler
    'Use a different workspace to prevent closing the default one _
        and thus screwing up other applications
    Set wsScriptsWorkspace = CreateWorkspace("ScriptBuilderWorkspace", "admin", "", dbUseJet)
    Workspaces.Append wsScriptsWorkspace
    Set dbScriptsDatabase = wsScriptsWorkspace.OpenDatabase(strScriptsDatabaseName, False, False)
    Call SaveSetting(App.Title, "Database", "Name", strScriptsDatabaseName)
Exit Sub
ConnectScriptsDatabase_ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub

Private Sub DisconnectScriptsDatabase()
'Created by Robin G Brown, 28/7/97
'This subroutine closes the database and returns without error
    On Error Resume Next
    'Close database and workspace
    dbScriptsDatabase.Close
    Set dbScriptsDatabase = Nothing
    wsScriptsWorkspace.Close
    Set wsScriptsWorkspace = Nothing
End Sub

Private Sub RefreshQueries()
'Created by Robin G Brown, 28/7/97
'Make sure all the data is up to date
'Sub Declares
    'Error Trap
    On Error Resume Next
    rsActionsQuery.Requery
    rsActions.Requery
End Sub

Property Get GetNumberOfParameters() As Integer
'Created by Robin G Brown, 28/7/97
'Return the number of parameters
'Function Declares
    'Error Trap
    On Error Resume Next
    Set rsParameters = dbScriptsDatabase.OpenRecordset(conParametersQuery, dbOpenDynaset)
    If Err.Number = 0 Then
        rsParameters.MoveLast
        GetNumberOfParameters = rsParameters.RecordCount
    Else
        GetNumberOfParameters = 0
    End If
End Property

Public Function GetParameter(ByVal lngParameterID As Long) As String
'Created by Robin G Brown, 23/7/97
'Return a string with the Parameter name
'Function Declares
    'Error Trap
    On Error Resume Next
    rsParameters.FindFirst "ParameterID = " & lngParameterID
    If rsParameters.NoMatch = True Then
        GetParameter = ""
    Else
        GetParameter = rsParameters.Fields("Value").Value
    End If
End Function

Public Function GetParameterID(ByVal strParameter As String) As Long
'Created by Robin G Brown, 23/7/97
'Return a string with the Parameter name
'Function Declares
    'Error Trap
    On Error Resume Next
    rsParameters.FindFirst "Parameter = '" & strParameter & "'"
    If rsParameters.NoMatch = True Then
        GetParameterID = 0
    Else
        GetParameterID = rsParameters.Fields("ParameterID").Value
    End If
End Function

Property Get GetNumberOfCommands() As Integer
'Created by Robin G Brown, 28/7/97
'Return the number of Commands
'Function Declares
    'Error Trap
    On Error Resume Next
    Set rsCommands = dbScriptsDatabase.OpenRecordset(conCommandsQuery, dbOpenDynaset)
    If Err.Number = 0 Then
        rsCommands.MoveLast
        GetNumberOfCommands = rsCommands.RecordCount
    Else
        GetNumberOfCommands = 0
    End If
End Property

Public Function GetCommand(ByVal lngCommandID As Long) As String
'Created by Robin G Brown, 23/7/97
'Return a string with the command name
'Function Declares
    'Error Trap
    On Error Resume Next
    rsCommands.FindFirst "CommandID = " & lngCommandID
    If rsCommands.NoMatch = True Then
        GetCommand = ""
    Else
        GetCommand = rsCommands.Fields("Command").Value
    End If
End Function

Property Get GetNumberOfServers() As Integer
'Created by Robin G Brown, 28/7/97
'Return the number of Servers
'Function Declares
    'Error Trap
    On Error Resume Next
    Set rsServers = dbScriptsDatabase.OpenRecordset(conServersQuery, dbOpenDynaset)
    If Err.Number = 0 Then
        rsServers.MoveLast
        GetNumberOfServers = rsServers.RecordCount
    Else
        GetNumberOfServers = 0
    End If
End Property

Public Function GetServer(ByVal lngServerNumber As Long) As String
'Created by Robin G Brown, 23/7/97
'Return a string with the Server name
'Function Declares
    'Error Trap
    On Error Resume Next
    rsServers.FindFirst "ServerID = " & lngServerNumber
    If rsServers.NoMatch = True Then
        GetServer = ""
    Else
        GetServer = rsServers.Fields("Server").Value
    End If
End Function

Property Get GetNumberOfActions() As Integer
'Created by Robin G Brown, 28/7/97
'Return the number of actions
'Function Declares
    'Error Trap
    On Error Resume Next
    rsActions.MoveLast
    GetNumberOfActions = rsActions.RecordCount
End Property

Public Function GetAction(ByVal lngActionID As Long) As String
'Created by Robin G Brown, 23/7/97
'Return a string with the action Name
'Function Declares
    'Error Trap
    On Error Resume Next
    rsActions.FindFirst "ActionID = " & lngActionID
    If rsActions.NoMatch = True Then
        GetAction = ""
    Else
        GetAction = rsActions.Fields("Name").Value
    End If
End Function

Public Function GetActionNumber(ByVal strActionName As String) As Long
'Created by Robin G Brown, 23/7/97
'Return a string with the action ID
'Function Declares
    'Error Trap
    On Error Resume Next
    rsActions.FindFirst "Name = '" & strActionName & "'"
    If rsActions.NoMatch = True Then
        GetActionNumber = 0
    Else
        GetActionNumber = rsActions.Fields("ActionID").Value
    End If
End Function

Public Function GetActionValue(ByVal lngActionID As Long) As String
'Created by Robin G Brown, 23/7/97
'Return a string with the action Value
'Function Declares
    'Error Trap
    On Error Resume Next
    rsActions.FindFirst "ActionID = " & lngActionID
    If rsActions.NoMatch = True Then
        GetActionValue = ""
    Else
        GetActionValue = "[" & rsActions.Fields("Value").Value & "]"
    End If
End Function

Public Function GetNumberOfActionParameters(ByVal lngActionID As Long) As String
'Created by Robin G Brown, 23/7/97
'Return a string with the action Value
'Function Declares
    'Error Trap
    On Error Resume Next
    rsActions.FindFirst "ActionID = " & lngActionID
    If rsActions.NoMatch = True Then
        GetNumberOfActionParameters = ""
    Else
        GetNumberOfActionParameters = rsActions.Fields("NumberOfParameters").Value
    End If
End Function

Public Function GetActionParameter(ByVal lngActionID As Long, ByVal bytParameterOrder As Byte) As String
'Created by Robin G Brown, 23/7/97
'Return argument description based on the action and the number of the argument
'Function Declares
Dim strCurrentAction As String
    'Error Trap
    On Error Resume Next
    rsActionsQuery.FindFirst "ActionID = " & lngActionID & " AND Order = " & bytParameterOrder
    If rsActionsQuery.NoMatch = True Then
        GetActionParameter = ""
    Else
        GetActionParameter = rsActionsQuery.Fields("Parameter").Value & "="
    End If
End Function

Public Function GetActionCommand(ByVal lngActionID As Long, ByVal bytParameterOrder As Byte) As String
'Created by Robin G Brown, 23/7/97
'Return argument command description based on the action and the number of the argument
'Function Declares
Dim strCurrentAction As String
    'Error Trap
    On Error Resume Next
    rsActionsQuery.FindFirst "ActionID = " & lngActionID & " AND Order = " & bytParameterOrder
    If rsActionsQuery.NoMatch = True Then
        GetActionCommand = ""
    Else
        GetActionCommand = rsActionsQuery.Fields("Command").Value
    End If
End Function

Public Function GetActionDefault(ByVal lngActionID As Long, ByVal bytParameterOrder As Byte) As String
'Created by Robin G Brown, 23/7/97
'Return argument command description based on the action and the number of the argument
'Function Declares
Dim strCurrentAction As String
    'Error Trap
    On Error Resume Next
    rsActionsQuery.FindFirst "ActionID = " & lngActionID & " AND Order = " & bytParameterOrder
    If rsActionsQuery.NoMatch = True Then
        GetActionDefault = ""
    Else
        GetActionDefault = rsActionsQuery.Fields("Default").Value
    End If
End Function

Public Function GetCommandFilter(ByVal strCommand As String) As String
'Created by Robin G Brown, 28/7/97
'Return the file filter (if applicable) for a given command
'Function Declares
    'Error Trap
    On Error Resume Next
    Set rsCommands = dbScriptsDatabase.OpenRecordset(conCommandsQuery, dbOpenDynaset)
    If Err.Number = 0 Then
        rsCommands.FindFirst "Command = '" & strCommand & "'"
        If rsActionsQuery.NoMatch = False Then
            GetCommandFilter = rsCommands.Fields("Filter").Value
            Exit Function
        End If
    End If
    GetCommandFilter = ""
End Function

Public Function AddAction(ByVal strActionName As String, ByVal strActionValue As String) As Long
'Created by Robin G Brown, 28/7/97
'Add a new action to the database, returns the ID of the new action _
    or 0 if there was an error
'Function Declares
Const conFunction = "AddAction"
Dim lngNextActionID As Long
    'Error Trap
    On Error GoTo AddAction_ErrorHandler
    With rsActions
        .MoveLast
        lngNextActionID = .Fields("ActionID").Value + 1
        .AddNew
        .Fields("ActionID").Value = lngNextActionID
        .Fields("Name").Value = strActionName
        .Fields("Value").Value = strActionValue
        .Fields("NumberOfParameters").Value = 0
        .Update
    End With
    AddAction = lngNextActionID
    'Refresh data
    Call RefreshQueries
Exit Function
'Error Handler
AddAction_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise vbObjectError, conModule & ":" & conFunction, "Unexpected error : " & Err.Description
    End Select
    AddAction = 0
    Exit Function
End Function

Public Function AddParameter(ByVal strParameterValue As String) As Long
'Created by Robin G Brown, 28/7/97
'Add a parameter, returns ParameterID of new Parameter, 0 on error
'Function Declares
Const conFunction = "AddParameter"
Dim rsParameter As Recordset
Dim lngNextParameterID As Long
    'Error Trap
    On Error GoTo AddParameter_ErrorHandler
    Set rsParameter = dbScriptsDatabase.OpenRecordset(conParametersQuery, dbOpenDynaset)
    With rsParameter
        .MoveLast
        lngNextParameterID = .Fields("ParameterID").Value + 1
        .AddNew
        .Fields("ParameterID").Value = lngNextParameterID
        .Fields("Parameter").Value = strParameterValue
        .Update
        .Close
    End With
    Set rsParameter = Nothing
    AddParameter = lngNextParameterID
Exit Function
'Error Handler
AddParameter_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise vbObjectError, conModule & ":" & conFunction, "Unexpected error : " & Err.Description
    End Select
    AddParameter = 0
    Exit Function
End Function

Public Function AddActionParameter(ByVal lngActionID As Long, ByVal lngParameterID As Long, ByVal bytOrder As Byte, ByVal strCommand As String, ByVal strDefault As String) As Long
'Created by Robin G Brown, 28/7/97
'Add a parameter to an action, returns true/false to indicate success
'Function Declares
Const conFunction = "AddActionParameter"
Dim rsActionParameter As Recordset
Dim lngNextParameterID As Long
    'Error Trap
    On Error GoTo AddActionParameter_ErrorHandler
    Set rsActionParameter = dbScriptsDatabase.OpenRecordset(conActionParameterQuery, dbOpenDynaset)
    With rsActionParameter
        .MoveLast
        .AddNew
        .Fields("ActionID").Value = lngActionID
        .Fields("ParameterID").Value = lngParameterID
        .Fields("Order").Value = bytOrder
        .Fields("Command").Value = strCommand
        .Fields("Default").Value = strDefault
        .Update
        .Close
    End With
    Set rsActionParameter = Nothing
    AddActionParameter = True
    'Refresh data
    Call RefreshQueries
Exit Function
'Error Handler
AddActionParameter_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise vbObjectError, conModule & ":" & conFunction, "Unexpected error : " & Err.Description
    End Select
    AddActionParameter = False
    Exit Function
End Function

Property Get DatabaseName() As String
'Created by Robin G Brown, 4/8/97
'Return the scripts database name and path
'Function Declares
    'Error Trap
    On Error Resume Next
    DatabaseName = strScriptsDatabaseName
End Property

Property Let DatabaseName(ByVal strNewDatabase As String)
'Created by Robin G Brown, 4/8/97
'Set the scripts database name and path
'Sub Declares
    'Error Trap
    On Error Resume Next
    strScriptsDatabaseName = strNewDatabase
End Property

Private Sub Class_Terminate()
'Created by Robin G Brown, 28/7/97
'Tidy up when the engine closes
'Sub Declares
    'Error Trap
    On Error Resume Next
    rsActions.Close
    rsActionsQuery.Close
    rsCommands.Close
    rsServers.Close
    rsParameters.Close
    Set rsParameters = Nothing
    Set rsServers = Nothing
    Set rsCommands = Nothing
    Set rsActions = Nothing
    Set rsActionsQuery = Nothing
    Call DisconnectScriptsDatabase
End Sub


