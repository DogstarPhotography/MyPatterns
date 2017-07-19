VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSVBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Created by Robin G Brown, 5th September 1997
'-----------------------------------------------------------------------------
'   This class module contains code for _
    a CSV database
'   Use cdb as a prefix for this class
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conClass = "CSVBase"
'Public properties
Public ctbTables As Collection
'Can't be arsed to set this up properly, so expose as a public variable, _
    should be set up using the 'House of Bricks' approach _
    or better yet the 'Custom House' approach _
    in Ch12 of Guide to the Perplexed, but needs additional tools to create
'Private properties
Private lngCounter As Long
Private lngReturn As Long
Private strDirectory As String

Private Sub Class_Initialize()
'Created by Robin G Brown, 5/9/97
'Sub Declares
    'Error Trap
    On Error Resume Next
    'Create a new collection
    Set ctbTables = New Collection
    'Set a default value for the path
    strDirectory = App.Path
End Sub

'-----------------------------------------------------------------------------
'   Public Functions and Methods
'-----------------------------------------------------------------------------

Public Sub AddTable(ByVal strTableName As String)
'Created by Robin G Brown, 5/9/97
'Add a new table to the collection
'Sub Declares
Const conSub = "AddTable"
Dim ctbNewTable As CSVTable
    'Error Trap
    On Error GoTo AddTable_ErrorHandler
    'Create a new table object
    Set ctbNewTable = New CSVTable
    'Read from the file into the object
    Call ctbNewTable.FileOpen(strDirectory & "\" & strTableName & ".CSV")
    'Add the object to the collection
    ctbTables.Add ctbNewTable, strTableName
    'Clear the reference
    Set ctbNewTable = Nothing
Exit Sub
'Error Handler
AddTable_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conClass & ":" & conSub, "Unexpected error : " & Err.Description
    End Select
    Exit Sub
End Sub

'-----------------------------------------------------------------------------
'   Properties
'-----------------------------------------------------------------------------

Property Get Directory() As String
'Created by Robin G Brown, 5/9/97
'Return the root directory of the CSV database
'Property Declares
Const conProperty = "Directory"
'Error Trap
    On Error Resume Next
    Directory = strDirectory
End Property

Property Let Directory(ByVal strNewDirectory As String)
'Created by Robin G Brown, 5/9/97
'Set the root directory of the CSV database
'Sub Declares
Const conProperty = "Directory"
    'Error Trap
    On Error Resume Next
    strDirectory = strNewDirectory
End Property

Private Sub Class_Terminate()
'Created by Robin G Brown, 5/9/97
'Sub Declares
Dim ctbCurrentTable As CSVTable
    'Error Trap
    On Error Resume Next
    'Save all of the tables
    For Each ctbCurrentTable In ctbTables
        If TypeOf ctbCurrentTable Is CSVTable Then
            ctbCurrentTable.FileSave
        End If
    Next
    'clear the reference
    Set ctbTables = Nothing
End Sub

