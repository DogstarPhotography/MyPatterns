Attribute VB_Name = "modMenuLib"
'Created by Robin G Brown, INSERT_LONG_DATE
'-----------------------------------------------------------------------------
'   This module contains code for _
    DESCRIPTION_OF_MODULE_FUNCTIONALITY_HERE
'-----------------------------------------------------------------------------
'   Special Instructions
'   Copy mnuImmediateError_Click or mnuMultiImmediateError_Click to Form modules
'   Uncomment routine and edit to call appropriate functions,
'   DelayedErrorMenuRoutine|ImmediateErrorMenuRoutine as required
'   This allows a common set of code to be created which can be called from multiple menus
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "modMenuLib"
Private lngCounter As Long
Private lngReturn As Long

'Public Sub mnuImmediateError_Click()
''Created by Robin G Brown, INSERT_DATE
''Call DelayedErrorMenuRoutine_OR_ImmediateErrorMenuRoutine in modMenuLib module
''Sub Declares
'Const conSub = "mnuImmediateError_Click"
'    'Error Trap
'    On Error GoTo mnuImmediateError_Click_ErrorHandler
'    'Call the approriate routine passing a reference to the current form
'    'Call modMenuLib.DelayedErrorMenuRoutine_OR_ImmediateErrorMenuRoutine(Me)
'Exit Sub
'mnuImmediateError_Click_ErrorHandler:
'    MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
'    Exit Sub
'End Sub

'Public Sub mnuMultiImmediateError_Click(Index As Integer)
''Created by Robin G Brown, INSERT_DATE
''Call appropriate DelayedErrorMenuRoutine_OR_ImmediateErrorMenuRoutine in modMenuLib module
''Sub Declares
'Const conSub = "mnuMultiImmediateError_Click"
'    'Error Trap
'    On Error GoTo mnuMultiImmediateError_Click_ErrorHandler
'    'Call the selected routine passing a reference to the current form
'    Select Case Index
''    Case 0
''        Call modMenuLib.DelayedErrorMenuRoutine_OR_ImmediateErrorMenuRoutine(Me)
'    Case Else
'        'Default - do nothing
'    End Select
'Exit Sub
'mnuMultiImmediateError_Click_ErrorHandler:
'    MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
'    Exit Sub
'End Sub

Public Sub DelayedErrorMenuRoutine(ByRef frmMenuSource As Form)
'Created by Robin G Brown, INSERT_DATE
'DESCRIPTION_OF_SUBROUTINE_FUNCTIONALITY_HERE
'Sub Declares
Const conSub = "DelayedErrorMenuRoutine"
    'Error Trap
    On Error Resume Next
    Err.Clear
    'CODE_HERE
    If Err.Number <> 0 Then
        'ERROR_HANDLING_CODE_HERE
        Err.Raise Err.Number, conModule & ":" & conSub, "Unexpected error : " & Err.Description
    End If
    'Clear the reference
    Set frmMenuSource = Nothing
End Sub

Public Sub ImmediateErrorMenuRoutine(ByRef frmMenuSource As Form)
'Created by Robin G Brown, INSERT_DATE
'DESCRIPTION_OF_SUBROUTINE_FUNCTIONALITY_HERE
'Sub Declares
Const conSub = "ImmediateErrorMenuRoutine"
    'Error Trap
    On Error GoTo ImmediateErrorMenuRoutine_ErrorHandler
    'CODE_HERE
    'Clear the reference
    Set frmMenuSource = Nothing
Exit Sub
'Error Handler
ImmediateErrorMenuRoutine_ErrorHandler:
    'Clear the reference
    Set frmMenuSource = Nothing
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conModule & ":" & conSub, "Unexpected error : " & Err.Description
    End Select
    Exit Sub
End Sub

