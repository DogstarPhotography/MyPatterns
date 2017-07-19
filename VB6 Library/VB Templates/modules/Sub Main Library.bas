Attribute VB_Name = "modMain"
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
Private Const conModule = "modMain"
Private lngCounter As Long
Private lngReturn As Long
Public Const conLoadSplashScreen = True
Public Const conUnloadSplashScreen = False

Sub Main()
'Created by Robin G Brown, INSERT_DATE
'DESCRIPTION_OF_SUBROUTINE_FUNCTIONALITY_HERE
'Sub Declares
Const conSub = "Main"
    'Error Trap
    On Error GoTo Main_ErrorHandler
    'Check for Instances of this program
    'Show splash form
    SplashScreen conLoadSplashScreen
    'Carry out any slow tasks
    'CODE_HERE
    'Show first form for startup
    frmBoilerPlate.Show
    'Get rid of splash form
    SplashScreen conUnloadSplashScreen
Exit Sub
'Error Handler
Main_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    'Exit the program
    Call ApplicationEnd
    Exit Sub
End Sub

Public Sub SplashScreen(ByVal booShow As Boolean)
'Created by Robin G Brown, 14/8/97
'Load/unload the splash screen
'Sub Declares
    'Error Trap
    On Error Resume Next
    If booShow = True Then
        frmSplash.Show
    Else
        Unload frmSplash
    End If
End Sub

Sub ApplicationEnd()
'Created by Robin G Brown, INSERT_DATE
'DESCRIPTION_OF_SUBROUTINE_FUNCTIONALITY_HERE
'Sub Declares
Const conSub = "ApplicationEnd"
    'Error Trap
    On Error Resume Next
    'CODE_HERE, _
        Close databases, etc _
        Free memory for objects _
        Save registry information _
        and Tidy up
    If Err.Number <> 0 Then
        'ERROR_HANDLING_CODE_HERE
        Err.Clear
    End If
    'End the application
    'End
End Sub

