Attribute VB_Name = "modUsefulLib"
'Created by Robin G Brown
'-----------------------------------------------------------------------------
'   Useful subroutines to include in any project
'   Many bits added from old General library, _
        search for OLDGENERALLIBRARY to find where they start
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "Useful Stuff"
Private intCounter As Integer
Private intReturn As Integer
'API declarations/constants
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const conFloat = 1
Global Const conSink = 0
'Sleep API call
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function FormInstances(ByVal strFormName As String) As Long
'Created by Robin G Brown, 20/8/97
'Check the forms collection to see how many copies of a form with this name are already loaded
'Function Declares
Const conFunction = "FormInstances"
Dim frmCurrentForm As Form
Dim lngInstances As Long
    'Error Trap
    On Error GoTo FormInstances_ErrorHandler
    lngInstances = 0
    For Each frmCurrentForm In Forms
        If frmCurrentForm.Name = strFormName Then
            lngInstances = lngInstances + 1
        End If
    Next
    FormInstances = lngInstances
Exit Function
'Error Handler
FormInstances_ErrorHandler:
    FormInstances = -1
    Exit Function
End Function

Public Sub Unavailable()
'Created by Robin G Brown, 25/4/97
'Unavailable message
'Sub Declares
    'Error Trap
    On Error Resume Next
    MsgBox "That item is currently unavailable.", vbOKOnly, App.ProductName
End Sub

Public Sub Pause()
'Created by Robin G Brown
'Wait for at least three seconds
'Sub Declares
    'Error Trap
    On Error Resume Next
    Call Sleep(3000)
End Sub

Public Sub PauseWithEvents()
'Created by Robin G Brown
'Wait for at least three seconds, calling doevents in the meantime
'Sub Declares
Dim varThen As Variant
    'Error Trap
    On Error Resume Next
    varThen = Now
    Do While DateDiff("s", varThen, Now) < 3
        'It's dirty, but you love it!
        DoEvents
    Loop
End Sub

Public Sub NoiseMaker()
'Created by Robin G Brown
'Make a noise to draw attention
    'Error Trap
    On Error Resume Next
    Beep
    Beep
    Beep
End Sub

Public Sub FloatWindow(ByRef X As Long, ByRef Action As Long)
'Float a Window so that it is 'Always on Top'
'When called by a form : _
    if action <> 0 makes the form Float (always on top) _
    if action = 0 'unFloats' the window
Dim wFlags As Long
Dim Result As Long
    On Error Resume Next
    wFlags = SWP_NOMOVE Or SWP_NOSIZE
    If Action <> 0 Then
        Result = SetWindowPos(X, HWND_TOPMOST, 0, 0, 0, 0, wFlags)
    Else
        Result = SetWindowPos(X, HWND_NOTOPMOST, 0, 0, 0, 0, wFlags)
    End If
End Sub

Public Function iMax(ByVal iValue1 As Integer, ByVal iValue2 As Integer) As Integer
' This function finds the maximum of 2 integer values
    On Error Resume Next
    If iValue1 >= iValue2 Then
        iMax = iValue1
    Else
        iMax = iValue2
    End If
End Function

Public Function iMin(ByVal iValue1 As Integer, ByVal iValue2 As Integer) As Integer
' This function finds the minimum of 2 integer values
    On Error Resume Next
    If iValue1 >= iValue2 Then
        iMin = iValue2
    Else
        iMin = iValue1
    End If
End Function


