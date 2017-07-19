Attribute VB_Name = "modTimeLib"
'Created by Robin G Brown
'-----------------------------------------------------------------------------
'   This module contains code for various time handling functions
'-----------------------------------------------------------------------------
'   Spoiler information
'   1 hour = 3600 seconds
'   1 day = 86400 seconds
'   1 week = 604800 seconds
'   A long integer has a maximum value of 2,147,483,647
'       which is 2,147,483 seconds
'       or 24 days
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "Time Handling Code"
'API Declarations
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long

Sub StarDate()
'Displays a Star Trek stardate
    On Error Resume Next
    MsgBox "StarDate : " & Format$(Now, "0000.00")
End Sub

Public Function GetDaysInMonth(ByVal intMonth As Integer) As Byte
'Created by Robin G Brown, 30/7/97
'Given a month, return the number of days in that month (inaccurate for leap years!)
'Function Declares
    'Error Trap
    On Error Resume Next
    Select Case intMonth
    Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12
        GetDaysInMonth = Choose(intMonth, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
    Case Else
        GetDaysInMonth = 30
    End Select
End Function

Public Function GetMonth(ByVal intMonth As Integer) As String
'Created by Robin G Brown, 21/4/97
'Returns a string representing the month
'Function Declares
    'Error Trap
    On Error Resume Next
    Select Case intMonth
    Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12
        GetMonth = Choose(intMonth, "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    Case Else
        GetMonth = ""
    End Select
End Function

Public Sub WaitFor(ByVal lngWaitSeconds As Long)
'Created by Robin G Brown, 22/7/97
'Wait for about lngWaitSeconds seconds, but will not wait for more than 1 day
'Sub Declares
Const conSub = "Pause"
Dim datStartTime As Date
    'Error Trap
    On Error GoTo WaitFor_ErrorHandler
    'WARNING! - no code will execute while the program is sleeping
    If lngWaitSeconds > 86400 Then
        Beep
        Exit Sub
    End If
    Call Sleep(lngWaitSeconds * 1000)
Exit Sub
'Error Handler
WaitFor_ErrorHandler:
    Exit Sub
End Sub

