VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
END
Attribute VB_Name = "clsTick2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Guide to Perplexed: Tick1
' Copyright (c) 1997 by Desaware Inc. All Rights Reserved
'
Option Explicit

Private DelayToUse As Long
Private CallbackObject As Object

Event DelayedCall()

' Callbackparam is an object that contains
' the method 'DelayedCall'
Public Sub TriggerDelay(Callbackparam As Object, delayval As Long)
   DelayToUse = delayval
   ' This is the object that we want to callback
   Set CallbackObject = Callbackparam
   StartTimer Me
End Sub

' Allow the module to access the delay value
' It's not public though
Friend Property Get Delay() As Long
   Delay = DelayToUse
End Property

' This is the timer expiration event from
' the module
Friend Sub TimerExpired()
   ' This is late bound
   If Not CallbackObject Is Nothing Then
      CallbackObject.DelayedCall
      ' Don't hold a reference to the object
      Set CallbackObject = Nothing
   Else
      RaiseEvent DelayedCall
   End If
End Sub

Private Sub Class_Initialize()
   Debug.Print "Object created"
End Sub

Private Sub Class_Terminate()
   Debug.Print "Object terminated"
End Sub
