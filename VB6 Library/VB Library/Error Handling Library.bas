Attribute VB_Name = "modErrorLib"
'Created by Robin G Brown
'-----------------------------------------------------------------------------
'   This module contains code for _
    handling errors
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "Error Handling Code"
Private intCounter As Integer
Private intReturn As Integer
Private strErrors() As String
'Error Codes
'User defined error codes run from 65535 backwards
'RGB error codes run from 50000 and up
Public Const errUndefined = 50000
Public Const errNoFileSelected = 50001

Public Sub NastyError(Optional ByVal varErrorNumber As Variant)
'Created by Robin G Brown
'Gives a nasty error!
'Sub Declares
Dim intVeryNasty As Integer
Dim intErrLen As Integer
Dim intStringCounter As Integer
Dim strErrMessage As String
Dim intReply As Integer
    'Error Trap
    On Error Resume Next
    If IsMissing(varErrorNumber) Then
        varErrorNumber = "RANDOM"
    End If
    If VarType(varErrorNumber) <> vbInteger Then
        'Create random integer
        'varErrorNumber = Int((65500 - 0 + 1) * Rnd + 0)
        varErrorNumber = Int((10 - 0 + 1) * Rnd + 0)
    End If
    strErrMessage = ""
    intVeryNasty = Int((100 - 1 + 1) * Rnd + 1)
    Select Case varErrorNumber
    Case 0
        strErrMessage = "Too many error messages, cannot continue."
    Case 1
        strErrMessage = "Please report for termination."
    Case 2
        strErrMessage = "WARNING : User intelligence below minimum required level : WARNING"
    Case 3
        strErrMessage = "...and another thing. Oh..."
    Case 4
        strErrMessage = "Please be quiet."
    Case 5
        strErrMessage = ""
    Case 6
        strErrMessage = ""
    Case 7
        strErrMessage = ""
    Case 8
        strErrMessage = ""
    Case 9
        strErrMessage = ""
    Case 10
        strErrMessage = ""
    Case Else
        intVeryNasty = 13
        'Generate a random string of garbled characters
        intErrLen = Int((50 - 5 + 1) * Rnd + 5)
        For intStringCounter = 1 To intErrLen
            strErrMessage = strErrMessage & Chr(Int((126 - 32 + 1) * Rnd + 32))
        Next
    End Select
    If intVeryNasty < 33 Then
        intReply = vbRetry
        Do While intReply = vbRetry
            intReply = MsgBox(strErrMessage, vbAbortRetryIgnore + vbDefaultButton2 + vbSystemModal + vbCritical, "!!!ERROR!!!ERROR!!!ERROR!!!")
            strErrMessage = ""
            intErrLen = Int((100 - 5 + 1) * Rnd + 5)
            For intStringCounter = 1 To intErrLen
                strErrMessage = strErrMessage & Chr(Int((126 - 32 + 1) * Rnd + 32))
            Next
        Loop
    Else
        MsgBox strErrMessage, vbOKCancel + vbExclamation, "Error Encountered"
    End If
    On Error GoTo 0
End Sub

Public Function UserDefinedError(ByRef intErrorNumber As Integer) As String
'Created by Robin G Brown, 19/2/97
'Returns a string depending on the input, _
    all error strings are present in this function
'Function Declares
    'Error Trap
    On Error Resume Next
    Select Case intErrorNumber
    Case errNoFileSelected
        UserDefinedError = "No File Selected"
    Case Else
        UserDefinedError = "Undefined error"
    End Select
    On Error GoTo 0
End Function

Public Function UDError(ByRef intErrorNumber As Integer) As String
'Created by Robin G Brown, 19/2/97
'Returns a string depending on the input, _
    reads from an array of errors created elsewhere
'Function Declares
    'Error Trap
    On Error Resume Next
    UDError = strErrors(intErrorNumber)
    On Error GoTo 0
End Function

Public Sub CreateErrorArray()
'Created by Robin G Brown, 19/2/97
'Creates an array of errors
'Sub Declares
    'Error Trap
    On Error Resume Next
    ReDim strErrors(1 To 1)
    On Error GoTo 0
End Sub


