VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
END
Attribute VB_Name = "Alarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Modified by Robin G Brown, 22nd July 1997
'-----------------------------------------------------------------------------
'   This class module contains code for _
    a simple callback object that can be used as an alarm:
'   SetAlarm - Use a callback method to raise an event which can be handled by the calling code
'      The calling object must have an AlarmRing 'event' in it (see below)
'   Use ala as a prefix for this class
'   Code copied and modified from Dan Appleman's Developing ActiveX Components with VB5 and the WIN32API
'
'Private WithEvents ALARMOBJECT As Alarm
'
'Public Sub ALARMOBJECT_AlarmRing()
''Created by Robin G Brown, INSERT_DATE
''This is the return 'event' from the Alarm object
''Sub Declares
'    'Error Trap
'    On Error Resume Next
'    'CODE_HERE
'End Sub
'
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "clsAlarm"
'Public properties
'Private properties
Private lngReturn As Long
Private lngDelayToUse As Long
Private objCallbackObject As Object
'Events
Event AlarmRing()

Public Sub SetAlarm(ByRef objCallbackParameter As Object, ByVal lngDelayValue As Long)
'Modified by Robin G Brown, 22/7/97
'objCallbackParameter is an object that contains the method 'Callback' (see above)
'Sub Declares
    'Error Trap
    On Error Resume Next
    'lngDelayValue is in seconds but lngDelayToUse is in milliseconds
    lngDelayToUse = lngDelayValue * 1000
    'This is the object that we want to callback
    Set objCallbackObject = objCallbackParameter
    Call StartTimer(Me)
End Sub

Friend Property Get Delay() As Long
'Modified by Robin G Brown, 22/7/97
'Allow the module to access the delay value, It's not public though
'Sub Declares
    'Error Trap
    On Error Resume Next
   Delay = lngDelayToUse
End Property

Friend Sub TimerExpired()
'Modified by Robin G Brown, 22/7/97
'This is the timer expiration event from the module
'Sub Declares
    'Error Trap
    On Error Resume Next
   ' This is late bound
   If Not objCallbackObject Is Nothing Then
      objCallbackObject.AlarmRing
      ' Don't hold a reference to the object
      Set objCallbackObject = Nothing
   Else
      RaiseEvent AlarmRing
   End If
End Sub

Public Sub CancelAlarm()
'Created by Robin G Brown, 22/7/97
'Cancel the alarm
'Sub Declares
    'Error Trap
    On Error Resume Next
    Call CancelTimer
End Sub

