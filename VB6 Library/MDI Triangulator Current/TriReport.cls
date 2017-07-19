VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TriReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Created by Robin G Brown, 2nd May 1997
'-----------------------------------------------------------------------------
'   This class module contains code for _
    DESCRIPTION_OF_CLASS_MODULE_FUNCTIONALITY_HERE
'   Use BPC as a prefix for this class
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "TriReport"
'Public properties
'Private properties
Private intCounter As Integer
Private intReturn As Integer
Public Sub ShowReport()
'Created by Robin G Brown, 2/5/97
'Show the main form
'Sub Declares
Const conSub = "ShowReport"
    'Error Trap
    On Error GoTo ShowReport_ErrorHandler
    frmGridReport.Show
Exit Sub
'Error Handler
ShowReport_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

