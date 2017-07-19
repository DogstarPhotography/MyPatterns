Attribute VB_Name = "modAlarmCallback"
'Modified by Robin G Brown, 22nd July 1997
'-----------------------------------------------------------------------------
'   This module contains code for _
    calling the Alarm object after a specified period
'   Code copied and modified from Dan Appleman's Developing ActiveX Components with VB5 and the WIN32API
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "modAlarmCallback"
' Timer identifier
Dim lngTimerID As Long
' Object for this timer
Dim alaTimerObject As Alarm
'API Declares
Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Public Sub StartTimer(ByRef alaCallingAlarm As Alarm)
'Modified by Robin G Brown, 22/7/97
'Start the timer
'Sub Declares
    'Error Trap
    On Error Resume Next
   Set alaTimerObject = alaCallingAlarm
   lngTimerID = SetTimer(0, 0, alaCallingAlarm.Delay, AddressOf TimerProc)
End Sub

Public Sub TimerProc(ByVal hwnd&, ByVal msg&, ByVal id&, ByVal currentime&)
'Modified by Robin G Brown, 22/7/97
'Callback function
'Sub Declares
    'Error Trap
    On Error Resume Next
   Call KillTimer(0, lngTimerID)
   lngTimerID = 0
   alaTimerObject.TimerExpired
   'Clear the object reference so it can delete
   Set alaTimerObject = Nothing
End Sub

Public Sub CancelTimer()
'Created by Robin G Brown, 22/7/97
'Cancel the timer
'Sub Declares
    'Error Trap
    On Error Resume Next
   Call KillTimer(0, lngTimerID)
   lngTimerID = 0
   'Clear the object reference so it can delete
   Set alaTimerObject = Nothing
End Sub

