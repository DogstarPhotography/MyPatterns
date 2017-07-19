Attribute VB_Name = "modMouseLib"
'Created by Robin G Brown, 23rd May 1997
'-----------------------------------------------------------------------------
'   This module contains code for _
    handling some mouse interaction, _
    and some boilerplate code for drag/drop operations
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "modMouseLib"
Private intCounter As Integer
Private intReturn As Integer
'Special Type for handling mice
Public Type MouseState
    'Use mst as a prefix
    x As Single
    y As Single
    Button As Integer
    Shift As Integer
End Type
Public mstCurrentMouseState As MouseState
Public booDragInProgress As Boolean

Public Function IsLeftDown(ByRef intButton As Integer) As Boolean
'Created by Robin G Brown, 27/5/97
'Return the state of the Shift key, True = Pressed, False otherwise
'Function Declares
Const conFunction = "IsLeftDown"
    'Error Trap
    On Error GoTo IsLeftDown_ErrorHandler
    IsLeftDown = (intButton And vbLeftButton) > 0
Exit Function
'Error Handler
IsLeftDown_ErrorHandler:
    IsLeftDown = False
    Exit Function
End Function

Public Function IsRightDown(ByRef intButton As Integer) As Boolean
'Created by Robin G Brown, 27/5/97
'Return the state of the Shift key, True = Pressed, False otherwise
'Function Declares
Const conFunction = "IsRightDown"
    'Error Trap
    On Error GoTo IsRightDown_ErrorHandler
    IsRightDown = (intButton And vbRightButton) > 0
Exit Function
'Error Handler
IsRightDown_ErrorHandler:
    IsRightDown = False
    Exit Function
End Function

Public Function IsMiddleDown(ByRef intButton As Integer) As Boolean
'Created by Robin G Brown, 27/5/97
'Return the state of the Shift key, True = Pressed, False otherwise
'Function Declares
Const conFunction = "IsMiddleDown"
    'Error Trap
    On Error GoTo IsMiddleDown_ErrorHandler
    IsMiddleDown = (intButton And vbMiddleButton) > 0
Exit Function
'Error Handler
IsMiddleDown_ErrorHandler:
    IsMiddleDown = False
    Exit Function
End Function

Public Function IsShiftDown(ByRef intShift As Integer) As Boolean
'Created by Robin G Brown, 27/5/97
'Return the state of the Shift key, True = Pressed, False otherwise
'Function Declares
Const conFunction = "IsShiftDown"
    'Error Trap
    On Error GoTo IsShiftDown_ErrorHandler
    IsShiftDown = (intShift And vbShiftMask) > 0
Exit Function
'Error Handler
IsShiftDown_ErrorHandler:
    IsShiftDown = False
    Exit Function
End Function

Public Function IsCtrlDown(ByRef intShift As Integer) As Boolean
'Created by Robin G Brown, 27/5/97
'Return the state of the Ctrl key, True = Pressed, False otherwise
'Function Declares
Const conFunction = "IsCtrlDown"
    'Error Trap
    On Error GoTo IsCtrlDown_ErrorHandler
    IsCtrlDown = (intShift And vbCtrlMask) > 0
Exit Function
'Error Handler
IsCtrlDown_ErrorHandler:
    IsCtrlDown = False
    Exit Function
End Function

Public Function IsAltDown(ByRef intShift As Integer) As Boolean
'Created by Robin G Brown, 27/5/97
'Return the state of the Alt key, True = Pressed, False otherwise
'Function Declares
Const conFunction = "IsAltDown"
    'Error Trap
    On Error GoTo IsAltDown_ErrorHandler
    IsAltDown = (intShift And vbAltMask) > 0
Exit Function
'Error Handler
IsAltDown_ErrorHandler:
    IsAltDown = False
    Exit Function
End Function

'-----------------------------------------------------------------------------
'   Some boilerplate code for drag/drop operations
'
'   Special instructions:
'   Add these routines to a module
'   Replace FORMORCONTROL with the name of the form or control
'   Get rid of the comments in front of the .Drag lines
'
'-----------------------------------------------------------------------------
Private Sub FORMORCONTROL_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'Created by Robin G Brown, 23/7/97
'Make sure that mstCurrentMouseState is set
'Sub Declares
    'Error Trap
    On Error Resume Next
    'First we need to find out where the mouse is _
        and what keys were pressed at the time
    mstCurrentMouseState.Button = Button
    mstCurrentMouseState.Shift = Shift
    mstCurrentMouseState.x = x
    mstCurrentMouseState.y = y
    'Then popup a menu if required
    If Button = vbRightButton Then
        'CODE_HERE
    End If
End Sub
Private Sub FORMORCONTROL_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'Created by Robin G Brown, 23/7/97
'Make sure that mstCurrentMouseState is set
'Sub Declares
    'Error Trap
    On Error Resume Next
    'First we need to find out where the mouse is _
        and what keys were pressed at the time
    mstCurrentMouseState.Button = Button
    mstCurrentMouseState.Shift = Shift
    mstCurrentMouseState.x = x
    mstCurrentMouseState.y = y
    'Cancel a drag/drop operation
    If booDragInProgress = True Then
        'FORMORCONTROL.Drag vbCancel
        booDragInProgress = False
    End If
End Sub

Private Sub FORMORCONTROL_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'Created by Robin G Brown, 23/7/97
'Start a drag/drop operation
'Sub Declares
    'Error Trap
    On Error Resume Next
    If IsLeftDown(Button) Then
        'FORMORCONTROL.Drag vbBeginDrag
        booDragInProgress = True
    End If
End Sub


