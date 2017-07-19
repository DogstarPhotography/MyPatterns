VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
END
Attribute VB_Name = "Sleeper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Modified by Robin G Brown, 22nd July 1997
'-----------------------------------------------------------------------------
'   This class module contains code for _
    a simple callback object that can be used as a sleep function
'   GoToSleep - Send the object to sleep, which waits for the time to expire before continuing
'   Use ala as a prefix for this class
'   Code copied and modified from Dan Appleman's Developing ActiveX Components with VB5 and the WIN32API
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "clsSleeper"
'Public properties
'Private properties
Private lngReturn As Long
Private Const ERROR_ALREADY_EXISTS = 183&
Private Const WAIT_OBJECT_0 = 0
Private Const WAIT_TIMEOUT = &H102&
Private lngTimerObjectHandle As Long
'Types
Private Type FILETIME
    'Use ftm as a prefix
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
'API Declares
Private Declare Function CreateWaitableTimer Lib "kernel32" Alias "CreateWaitableTimerA" (ByVal lpSemaphoreAttributes As Long, ByVal bManualReset As Long, ByVal lpName As String) As Long
Private Declare Function OpenWaitableTimer Lib "kernel32" Alias "OpenWaitableTimerA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Private Declare Function SetWaitableTimer Lib "kernel32" (ByVal hTimer As Long, lpDueTime As FILETIME, ByVal lPeriod As Long, ByVal pfnCompletionRoutine As Long, ByVal lpArgToCompletionRoutine As Long, ByVal fResume As Long) As Long
Private Declare Function CancelWaitableTimer Lib "kernel32" (ByVal hTimer As Long)
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Sub Class_Initialize()
'Modified by Robin G Brown, 22/7/97
'Initialize the EXE
'Sub Declares
Dim ftmSetTime As FILETIME
    'Error Trap
    On Error Resume Next
   lngTimerObjectHandle = CreateWaitableTimer(0, True, "AlarmTimerAlarm")
   If Err.LastDllError = ERROR_ALREADY_EXISTS Then
      ' Another app already Modified the timer,
      ' we may need to go to sleep now
      Err.Raise vbObjectError, conModule, "Cannot create Timer"
   Else
      ftmSetTime.dwLowDateTime = -1
      ftmSetTime.dwHighDateTime = -1
      lngReturn = SetWaitableTimer(lngTimerObjectHandle, ftmSetTime, 0, 0, 0, 0)
   End If
End Sub

Public Sub GoToSleep(ByVal lngSeconds As Long)
'Modified by Robin G Brown, 22/7/97
'Function to go to sleep
'Sub Declares
Dim ftmSetTime As FILETIME
Dim lngIsBusy As Long
Dim dblDelay As Double
Dim dblDelayLow As Double
Dim dblUnit32 As Double
    'Error Trap
    On Error Resume Next
   ' Is object signaled now?
   lngIsBusy = WaitForSingleObject(lngTimerObjectHandle, 0)
   ' We know WAIT_OBJECT_0 is 0, so False indicates signaled
   If Not lngIsBusy Then
      ' Set the timer
      dblUnit32 = CDbl(&H10000) * CDbl(&H10000)
      dblDelay = CDbl(lngSeconds)
      dblDelay = dblDelay * 1000 * 10000
      ' dblDelay is now 100ns
      ftmSetTime.dwHighDateTime = -CLng(dblDelay / dblUnit32) - 1
      dblDelayLow = -dblUnit32 * (dblDelay / dblUnit32 - Fix(dblDelay / dblUnit32))
      If dblDelayLow < CDbl(&H80000000) Then
         dblDelayLow = dblUnit32 + dblDelayLow
      End If
      ftmSetTime.dwLowDateTime = CLng(dblDelayLow)
      lngReturn = SetWaitableTimer(lngTimerObjectHandle, ftmSetTime, 0, 0, 0, 0)
   End If
   'PRE_WAIT_CODE
   lngIsBusy = WaitForSingleObject(lngTimerObjectHandle, -1)
   'POST_WAIT_CODE
End Sub

Private Sub Class_Terminate()
'Modified by Robin G Brown, 22/7/97
'Terminate the EXE
'Sub Declares
    'Error Trap
    On Error Resume Next
    'Close the handle on unload
    Call CloseHandle(lngTimerObjectHandle)
End Sub
