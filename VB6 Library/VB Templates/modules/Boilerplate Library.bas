Attribute VB_Name = "modBoilerLib"
'Created by Robin G Brown, INSERT_LONG_DATE
'-----------------------------------------------------------------------------
'   This module contains code for _
    DESCRIPTION_OF_MODULE_FUNCTIONALITY_HERE
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "modBoilerLib"
Private lngCounter As Long
Private lngReturn As Long

Public Sub VerySimpleSubOrMenu()
'Created by Robin G Brown, INSERT_DATE
'DESCRIPTION_OF_SUBROUTINE_FUNCTIONALITY_HERE
    'Error Trap
    On Error Resume Next
    'CODE_HERE
End Sub

Public Sub DelayedErrorSub()
'Created by Robin G Brown, INSERT_DATE
'DESCRIPTION_OF_SUBROUTINE_FUNCTIONALITY_HERE
'Sub Declares
    'Error Trap
    On Error Resume Next
    Err.Clear
    'CODE_HERE
    If Err.Number <> 0 Then
        'ERROR_HANDLING_CODE_HERE
    End If
End Sub

Public Function DelayedErrorFunction() As Long
'Created by Robin G Brown, INSERT_DATE
'DESCRIPTION_OF_FUNCTIONALITY_HERE
'Function Declares
    'Error Trap
    On Error Resume Next
    DelayedErrorFunction = False
    Err.Clear
    'CODE_HERE
    If Err.Number <> 0 Then
        'ERROR_HANDLING_CODE_HERE
    End If
    DelayedErrorFunction = True
End Function

Public Sub ImmediateErrorSub()
'Created by Robin G Brown, INSERT_DATE
'DESCRIPTION_OF_SUBROUTINE_FUNCTIONALITY_HERE
'Sub Declares
Const conSub = "ImmediateErrorSub"
    'Error Trap
    On Error GoTo ImmediateErrorSub_ErrorHandler
    'CODE_HERE
Exit Sub
'Error Handler
ImmediateErrorSub_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Public Function ImmediateErrorFunction() As Long
'Created by Robin G Brown, INSERT_DATE
'DESCRIPTION_OF_FUNCTIONALITY_HERE
'Function Declares
Const conFunction = "ImmediateErrorFunction"
    'Error Trap
    On Error GoTo ImmediateErrorFunction_ErrorHandler
    'CODE_HERE
    ImmediateErrorFunction = True
Exit Function
'Error Handler
ImmediateErrorFunction_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conFunction
    End Select
    ImmediateErrorFunction = False
    Exit Function
End Function
