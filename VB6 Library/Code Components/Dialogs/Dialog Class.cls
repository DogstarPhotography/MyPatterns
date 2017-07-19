VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Created by Robin G Brown, 5th August 1997
'-----------------------------------------------------------------------------
'   This class module contains code for _
    showing standard dialogs
'   Use sdl as a prefix for this class
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "Dialog"
'Public properties
'Private properties
Private lngCounter As Long
Private lngReturn As Long
'Dialog properties
Private dteDateTimeProperty As Date
Private strDatabaseNameProperty As String

Private Sub Class_Initialize()
'Created by Robin G Brown, 5/8/97
'Set default values
'Sub Declares
    'Error Trap
    On Error Resume Next
    dteDateTimeProperty = 0
    strDatabaseNameProperty = ""
End Sub

Public Sub ShowDateTime()
'Created by Robin G Brown, 5/8/97
'Show the DateTime dialog to set a value for dteDateTimeProperty
'Sub Declares
Const conSub = "ShowDateTime"
    'Error Trap
    On Error GoTo ShowDateTime_ErrorHandler
    Load frmDateTime
    With frmDateTime
        .ShowDateAndTime
        .DateTime = dteDateTimeProperty
        .Show vbModal
        dteDateTimeProperty = .DateTime
    End With
    Unload frmDateTime
Exit Sub
'Error Handler
ShowDateTime_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conModule, "Unexpected error : " & Err.Description
    End Select
    Exit Sub
End Sub

Public Sub ShowDate()
'Created by Robin G Brown, 5/8/97
'Show the DateTime dialog to set a value for dteDateTimeProperty
'Sub Declares
Const conSub = "ShowDate"
    'Error Trap
    On Error GoTo ShowDate_ErrorHandler
    Load frmDateTime
    With frmDateTime
        .ShowDateOnly
        .DateTime = dteDateTimeProperty
        .Show vbModal
        dteDateTimeProperty = .DateTime
    End With
    Unload frmDateTime
Exit Sub
'Error Handler
ShowDate_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conModule, "Unexpected error : " & Err.Description
    End Select
    Exit Sub
End Sub

Public Sub ShowTime()
'Created by Robin G Brown, 5/8/97
'Show the DateTime dialog to set a value for dteDateTimeProperty
'Sub Declares
Const conSub = "ShowTime"
    'Error Trap
    On Error GoTo ShowTime_ErrorHandler
    Load frmDateTime
    With frmDateTime
        .ShowTimeOnly
        .DateTime = dteDateTimeProperty
        .Show vbModal
        dteDateTimeProperty = .DateTime
    End With
    Unload frmDateTime
Exit Sub
'Error Handler
ShowTime_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conModule, "Unexpected error : " & Err.Description
    End Select
    Exit Sub
End Sub

Property Get DateTime() As Date
'Created by Robin G Brown, 5/8/97
'Function Declares
    'Error Trap
    On Error Resume Next
    DateTime = dteDateTimeProperty
End Property

Property Let DateTime(ByVal NewDateTime As Date)
'Created by Robin G Brown, 5/8/97
'Sub Declares
    'Error Trap
    On Error Resume Next
    dteDateTimeProperty = NewDateTime
End Property

Public Sub ShowSelectDatabase()
'Created by Robin G Brown, 5/8/97
'Show the select database dialog
'Sub Declares
Const conSub = "ShowSelectDatabase"
    'Error Trap
    On Error GoTo ShowSelectDatabase_ErrorHandler
    Load frmSelectDatabase
    With frmSelectDatabase
        .DatabaseName = strDatabaseNameProperty
        .Show vbModal
        strDatabaseNameProperty = .DatabaseName
    End With
    Unload frmSelectDatabase
Exit Sub
'Error Handler
ShowSelectDatabase_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conModule, "Unexpected error : " & Err.Description
    End Select
    Exit Sub
End Sub

Property Get DatabaseName() As String
'Created by Robin G Brown, 5/8/97
'Function Declares
    'Error Trap
    On Error Resume Next
    DatabaseName = strDatabaseNameProperty
End Property

Property Let DatabaseName(ByVal NewDatabaseName As String)
'Created by Robin G Brown, 5/8/97
'Sub Declares
    'Error Trap
    On Error Resume Next
    strDatabaseNameProperty = NewDatabaseName
End Property



