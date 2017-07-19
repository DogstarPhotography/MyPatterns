Attribute VB_Name = "modDatabaseLibrary"
'Created by Robin G Brown, a long time ago.
'-----------------------------------------------------------------------------
'   This module contains code for _
    handling database functions
'-----------------------------------------------------------------------------
'   Minor problem with this library :
'   Ther is only one workspace created with this library so you must be careful
'       when opening and closing databases not to also close the workspace accidentally
'   Optional parameters to the Connect and Disconnect functions can be set
'       to prevent the main workspace being closed
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "modDatabaseLibrary"
Private intCounter As Integer
Private intReturn As Integer
'Public Variables for ConnectSQLDatabase
Public gstrSQLDBName As String
Public gstrUserID As String
Public gstrPassword As String
Public gstrServer As String
Public dbSQLDataBase As Database
Public wsMainWorkSpace As Workspace
'Public Variables for ConnectDatabase
Public dbDataBase As Database
Public gstrDBName As String

Public Function ConnectSQLDatabase(Optional varConnectWorkspaceFlag As Variant = True) As Integer
'Created by Robin G Brown, 30/10/95
'This function is designed to connect to an SQL database using several global parameters _
    declared in module declarations section
'Returns true if succesful, false if an error
Dim strConnect As String
    On Error GoTo ConnectSQLDatabase_ErrorHandler
    ConnectSQLDatabase = False
    If varConnectWorkspaceFlag = True Then
        Set wsMainWorkSpace = Workspaces(0)
    End If
    strConnect = "ODBC" & _
        ";DSN=" & gstrSQLDBName & _
        ";UID=" & gstrUserID & _
        ";PWD=" & gstrPassword & _
        ";APP=" & App.Title & _
        ";WSID=" & gstrServer & _
        ";DATABASE=" & gstrSQLDBName
    Set dbSQLDataBase = wsMainWorkSpace.OpenDatabase("", False, False, strConnect)
    ConnectSQLDatabase = True
Exit Function
ConnectSQLDatabase_ErrorHandler:
    ConnectSQLDatabase = False
    Exit Function
End Function

Public Function CompactThisDatabase(strDBName As String, strDBNewName As String) As Integer
'By Mark M, 2/11/95, edited by RGB/17/1/96
'The database and all dependant recordsets/ DAO Objects must be closed, _
    AND AVAILABLE FOR EXCLUSIVE OPENING
'This function will try to Compact a given database to the new file name.
'The file names nust include the FULL PATH including network parameters if any.
    On Error GoTo CompactThisDatabase_Errorhandler
    CompactThisDatabase = False
    DBEngine.CompactDatabase strDBName, strDBNewName
    CompactThisDatabase = True
Exit Function
CompactThisDatabase_Errorhandler:
    CompactThisDatabase = False
    Exit Function
End Function

Public Sub DisconnectSQLDatabase(Optional varCloseWorkspaceFlag As Variant = True)
'Created by Robin G Brown, 17/1/96
'This subroutine closes the SQL database and returns without error
    On Error Resume Next
    'Close database and workspace
    dbSQLDataBase.Close
    Set dbSQLDataBase = Nothing
    If varCloseWorkspaceFlag = True Then
        wsMainWorkSpace.Close
        Set wsMainWorkSpace = Nothing
    End If
    'gintServerConnection = False
    On Error GoTo 0
End Sub

Public Sub CloseWorkspace()
'By Mark M, 30/10/95
'Closes the Workspace object separate to any Database objects
    On Error Resume Next
    wsMainWorkSpace.Close
    Set wsMainWorkSpace = Nothing
    On Error GoTo 0
End Sub

Public Function CopyTableDefIndexes(ByRef tbldefSource As TableDef, ByRef tbldefTarget As TableDef) As Integer
'By Mark M 9/Nov/95, edited by RGB/17/1/96

'NOT TESTED!!!

'This should loop through all indexes in a given Table def
'And copy them to a Target table def
'It uses the CreateTableDefIndex function.
Dim idxNewIndex As Index
Dim idxSourceIndex As Index
Dim strIndexName As String
Dim strFieldName As String
Dim intResult As Integer
Dim fldNewField As Field
Dim fldSourceField As Field
    On Error GoTo CopyTableDefIndexes_ErrorHandler
    CopyTableDefIndexes = False
    Screen.MousePointer = 11 'hourglass
    'Set intSourceIndex = dbDBForNewTable.CreateTableDef(strNewTableName)
    For Each idxSourceIndex In tbldefSource.Indexes
    'set and append copies of all dynaBatchEdit fields
    'For intCounter = 0 To tbldefSource.Indexes.Count - 1
        'set main parameters to variables:
        strIndexName = idxSourceIndex.Name
        'intPrimary = 0
        'intUnique = 0
        'intIgnoreNulls = 0
        'intRequired = 0
        'args are: (ToUseTableDef,strIndexname,intSetPrimary,intSetRequired,intSetUnique,intSetIgnoreNulls})
        'this will duplicate the properties
        intResult = CreateIndexOnTableDef(tbldefTarget, idxNewIndex, idxSourceIndex.Name, idxSourceIndex.Primary, idxSourceIndex.Required, idxSourceIndex.Unique, idxSourceIndex.IgnoreNulls)
        If intResult <> True Then GoTo CopyTableDefIndexes_FailedToCreate
        'append field objects:
        For Each fldSourceField In idxSourceIndex.Fields
            strFieldName = fldSourceField.Name
            Set fldNewField = idxNewIndex.CreateField(strFieldName)
            'append field
            idxNewIndex.Fields.Append fldNewField
        Next fldSourceField
        'LAST, append new index
        tbldefTarget.Indexes.Append idxNewIndex
CopyTableDefIndexes_Continue:
    Next idxSourceIndex
    CopyTableDefIndexes = True
    Screen.MousePointer = 0 'default
Exit Function
CopyTableDefIndexes_ErrorHandler:
    CopyTableDefIndexes = False
    Screen.MousePointer = 0
    Exit Function
CopyTableDefIndexes_FailedToCreate:
    'intResult = IDYES ' MsgBox("Unexpected Error - " & Err & "; " & Error & Chr$(10) & "Unable to create Index >" & strIndexName & "<." & Chr$(10) & "Continue Appending Indexes for this Table?", MB_ICONQUESTION + MB_YESNO + MB_DEFBUTTON2, "SmithKline Beecham")
    'If intResult = IDYES Then
    If 1 Then
        If Err <> 0 Then
            Resume Next
        Else
            GoTo CopyTableDefIndexes_Continue
        End If
    Else
        Screen.MousePointer = 0
        Exit Function
    End If
End Function

Public Function DeleteIndexOnTableDef(ByRef dbDBToUse As Database, ByVal strTableDefName As String, ByVal strIndexName As String) As Integer
'By Mark M, 9/Nov/95, edited by RGB/17/1/96

'To DELETE the given index on the given table def.

'NOTE: UNTESTED AS OF 9 NOV 95

    On Error GoTo DeleteIndexOnTableDef_Errorhandler
    DeleteIndexOnTableDef = True
    dbDBToUse.TableDefs(strTableDefName).Indexes.Delete strIndexName

Exit Function
DeleteIndexOnTableDef_Errorhandler:
    DeleteIndexOnTableDef = False
    'MsgBox "Unexpected Error - Unable to Delete Index >" & strIndexName & "<.", MB_ICONINFORMATION, "SmithKline Beecham"
    Exit Function

End Function

Public Function CreateIndexOnTableDef(ByRef tbldefTableToUse As TableDef, idxNewIndex As Index, strIndexName As String, Optional ByVal intSetRequired As Variant, Optional ByVal intSetPrimary As Variant, Optional ByVal intSetUnique As Variant, Optional ByVal intSetIgnoreNulls As Variant) As Integer
'By Mark M, 9/Nov/95, edited by RGB/17/1/96

'NOTE: UNTESTED AS OF 9 NOV.

'args are: (ToUseTableDef,strIndexname,intSetPrimary,intSetRequired,intSetUnique,intSetIgnoreNulls})
'The last four index property arguments are optional, and should be True/False.
'The default for each is False
'All relevant Fields must also be appended to the index
'To set the given index on the given table def, with the arguments as parameters.

Dim tbldefTableToModify As TableDef
'Dim idxNewIndex As Index

    On Error GoTo CreateIndexOnTableDef_Errorhandler
    CreateIndexOnTableDef = False
    If IsMissing(intSetPrimary) Then intSetPrimary = False
    If IsMissing(intSetRequired) Then intSetRequired = False
    If IsMissing(intSetUnique) Then intSetUnique = False
    If IsMissing(intSetIgnoreNulls) Then intSetIgnoreNulls = False
    If intSetPrimary <> True And intSetPrimary <> False Then intSetPrimary = False
    If intSetRequired <> True And intSetRequired <> False Then intSetRequired = False
    If intSetUnique <> True And intSetUnique <> False Then intSetUnique = False
    If intSetIgnoreNulls <> True And intSetIgnoreNulls <> False Then intSetIgnoreNulls = False
    Set idxNewIndex = tbldefTableToUse.CreateIndex(strIndexName)
    idxNewIndex.Primary = intSetPrimary
    idxNewIndex.Required = intSetRequired
    idxNewIndex.Unique = intSetUnique
    idxNewIndex.IgnoreNulls = intSetIgnoreNulls
    tbldefTableToUse.Indexes.Append idxNewIndex
    CreateIndexOnTableDef = True

Exit Function
CreateIndexOnTableDef_Errorhandler:
    CreateIndexOnTableDef = False
    'MsgBox "Unexpected Error - Unable to Create New Index >" & strIndexName & "<", , "SmithKline Beecham"
    Exit Function

End Function

Public Function SetWorkspace() As Boolean
'By Mark M, 30/10/95, edited by RGB/8/7/97
'Sets the Workspace object separate to any Database
'objects
    On Error GoTo SetWorkspace_ErrorHandler
    SetWorkspace = False
    Set wsMainWorkSpace = Workspaces(0)
    SetWorkspace = True
Exit Function
SetWorkspace_ErrorHandler:
    SetWorkspace = False
    Exit Function
End Function

Public Function WriteCSVFileFromRecordset(ByRef dynaSource As Recordset, ByVal strCSVFileName As String) As Integer
'Created by Robin G Brown, 27/4/95
'This function writes the contents of a recordset to a CSV file _
    return value = number of lines written or false if there is an error
Dim intFieldCounter As Integer
Dim intCSVFileNumber As Integer
Dim intLineCounter As Integer
    On Error GoTo WriteCSVFileFromRecordset_ErrorHandler
    intLineCounter = 0
    WriteCSVFileFromRecordset = intLineCounter
    'Move to start of dynaset and get file number
    dynaSource.MoveFirst
    intCSVFileNumber = FreeFile
    'Open file
    Open strCSVFileName For Output As intCSVFileNumber
    'Write column titles
    For intFieldCounter = 0 To dynaSource.Fields.Count - 1
        Print #intCSVFileNumber, dynaSource.Fields(intFieldCounter).Name; ",";
    Next
    'Add LF & CR
    Print #intCSVFileNumber,
    'Increment counter
    intLineCounter = intLineCounter + 1
    'Loop through records
    Do While Not dynaSource.EOF
        For intFieldCounter = 0 To dynaSource.Fields.Count - 1
            If IsNull(dynaSource.Fields(intFieldCounter).Value) Then
                Print #intCSVFileNumber, "NULL,";
            Else
                Print #intCSVFileNumber, dynaSource.Fields(intFieldCounter).Value; ",";
            End If
        Next
        Print #intCSVFileNumber,
        intLineCounter = intLineCounter + 1
        dynaSource.MoveNext
    Loop
    'Close file
    Close intCSVFileNumber
    'Return number of lines written
    WriteCSVFileFromRecordset = intLineCounter
Exit Function
WriteCSVFileFromRecordset_ErrorHandler:
    'MsgBox "" & Err & ": " & Error, , ""
    On Error Resume Next
    Kill strCSVFileName
    WriteCSVFileFromRecordset = False
    On Error GoTo 0
    Exit Function
End Function

Public Sub CSVStringToVariantArray(ByVal strCSVData As String, ByRef varTargetArray() As Variant)
'Created by Robin G Brown, unknown date
'This sub fills the first dimension of a variant array with values from a CSV string
Dim intCurrentPosition As Integer
Dim intNextPosition As Integer
Dim intArraycounter As Integer
    On Error GoTo CSVStringToVariantArray_ErrorHandler
    intCurrentPosition = 1
    intNextPosition = 1
    For intArraycounter = LBound(varTargetArray, 1) To UBound(varTargetArray, 1)
        intNextPosition = InStr(intCurrentPosition, strCSVData, ",")
        If intNextPosition <> 0 Then
            varTargetArray(intArraycounter) = Mid$(strCSVData, intCurrentPosition, (intNextPosition - intCurrentPosition))
            intCurrentPosition = intNextPosition + 1
        Else
            Exit Sub
        End If
    Next
Exit Sub
CSVStringToVariantArray_ErrorHandler:
    Exit Sub
End Sub

Public Function RepairThisDatabase(strDBName As String) As Integer
'By Mark M, 2/11/95, edited by RGB/17/1/96
'The database and all dependant recordsets/ DAO Objects must be closed,
'AND AVAILABLE FOR EXCLUSIVE OPENING
'This function will try to Repair a given database.
'The file name nust include the FULL PATH including network parameters if any.
    
    On Error GoTo RepairThisDatabase_Errorhandler
    RepairThisDatabase = False
    DBEngine.RepairDatabase strDBName
    RepairThisDatabase = True
    
Exit Function
RepairThisDatabase_Errorhandler:
    RepairThisDatabase = False
    Exit Function

End Function

Public Function ConnectDatabase(Optional varConnectWorkspaceFlag As Variant = True) As Integer
'Created by Robin G Brown, 17/1/96
'This function is designed to connect to an Access database using several global parameters _
    which are declared in the module declarations section
'Returns true if succesful, false if an error
    'gstrDBName must be a full path and file name
Dim strConnect As String
    'Error trap
    On Error GoTo ConnectDatabase_ErrorHandler
    ConnectDatabase = False
    If varConnectWorkspaceFlag = True Then
        Set wsMainWorkSpace = Workspaces(0)
    End If
    strConnect = ""
    Set dbDataBase = wsMainWorkSpace.OpenDatabase(gstrDBName, False, False, strConnect)
    ConnectDatabase = True
Exit Function
ConnectDatabase_ErrorHandler:
    ConnectDatabase = False
    Exit Function
End Function

Public Sub DisconnectDatabase(Optional varCloseWorkspaceFlag As Variant = True)
'Created by Robin G Brown, 17/1/96
'This subroutine closes the database and returns without error
    On Error Resume Next
    'Close database and workspace
    dbDataBase.Close
    Set dbDataBase = Nothing
    If varCloseWorkspaceFlag = True Then
        wsMainWorkSpace.Close
        Set wsMainWorkSpace = Nothing
    End If
End Sub

Public Sub ConnectSelectedDatabase(Optional ByVal varDefaultDatabaseName As Variant)
'Created by Robin G Brown, 15/9/97
'Carry out any tasks required to set up for using a database
'Function Declares
Const conFunction = "ConnectSelectedDatabase"
Dim strDefaultDatabaseName As String
Dim sdldialog As Dialog
Dim lngReturn As Long
    'Error Trap
    On Error GoTo ConnectSelectedDatabase_ErrorHandler
    If IsMissing(varDefaultDatabaseName) Then
        strDefaultDatabaseName = CStr(varDefaultDatabaseName)
    End If
    'Read any registry information to set database name
    gstrDBName = GetSetting(App.Title, "Database", "Name", strDefaultDatabaseName)
    'set the workspace first
    Call SetWorkspace
    Set sdldialog = New Dialog
    'Connect without setting the workspace
    lngReturn = ConnectDatabase(False)
    'If the connect failed, get the user to select a database to connect to
    Do While lngReturn = False
        lngReturn = MsgBox("Could not connect to database, this program may not run correctly, click OK to select a database, Cancel to exit.", vbOKCancel, "Database set up")
        If lngReturn = vbOK Then
            'Show a 'Standard' dialog for selecting a database
            With sdldialog
                .DatabaseName = strDefaultDatabaseName
                .ShowSelectDatabase
                gstrDBName = .DatabaseName
            End With
        Else
            Err.Raise vbObjectError, App.Title, "Database not found."
        End If
        'Attempt to connect the database again
        lngReturn = ConnectDatabase(False)
    Loop
    Set sdldialog = Nothing
    'Save the database setting away for next time
    Call SaveSetting(App.Title, "Database", "Name", gstrDBName)
Exit Sub
'Error Handler
ConnectSelectedDatabase_ErrorHandler:
    Set sdldialog = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub


