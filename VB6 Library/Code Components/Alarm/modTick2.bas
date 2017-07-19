Attribute VB_Name = "modTick2"
' Guide to Perplexed: Tick1
' Copyright (c) 1997 by Desaware Inc. All Rights Reserved

Option Explicit

' Timer identifier
Dim TimerID&

' Object for this timer
Dim TimerObject As clsTick2

Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long


Public Sub StartTimer(callingobject As clsTick2)
   Set TimerObject = callingobject
   TimerID = SetTimer(0, 0, callingobject.Delay, AddressOf TimerProc)
End Sub

' Callback function
Public Sub TimerProc(ByVal hwnd&, ByVal msg&, ByVal id&, ByVal currentime&)
   Call KillTimer(0, TimerID)
   TimerID = 0
   TimerObject.TimerExpired
   ' And clear the object reference so it can delete
   Set TimerObject = Nothing
End Sub
