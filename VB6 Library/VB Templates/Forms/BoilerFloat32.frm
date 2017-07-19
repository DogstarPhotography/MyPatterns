VERSION 5.00
Begin VB.Form frmBoilerPlate 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Boiler Plate"
   ClientHeight    =   4455
   ClientLeft      =   1155
   ClientTop       =   1560
   ClientWidth     =   7305
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4455
   ScaleWidth      =   7305
End
Attribute VB_Name = "frmBoilerPlate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Robin G Brown, INSERT_LONG_DATE
'-----------------------------------------------------------------------------
'   This form contains code for _
    DESCRIPTION_OF_FORM_FUNCTIONALITY_HERE
'-----------------------------------------------------------------------------
'   SPECIAL_INSTRUCTIONS
'   1. Fill in DESCRIPTION_OF_FORM_FUNCTIONALITY_HERE above
'   2. Replace all INSERT_DATE with date of creation
'   4. Rename form in Properties window (F4)
'   5. Replace all frmBoilerPlate with name of form
'   6. Fill in all DEFAULT_BEHAVIOUR
'   7. Make sure the FloatWindow function from the Useful Stuff library is included
'   8. Delete this SPECIAL_INSTRUCTIONS set of comments
'   9. Save File As... with suitable filename
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "frmBoilerPlate"
Private intCounter As Integer
Private intReturn As Integer
'Declare constants for boilerplate code
Private Const conCentering = True    'causes form to center on screen when shown, check form_resize event
Private Const conFloating = False    'causes form to be 'always on top', check form_activate and form_unload events for other bits
Private Sub Form_Activate()
'Created by Robin G Brown, INSERT_DATE
'DEFAULT_BEHAVIOUR
'Sub Declares
Const conSub = "Form_Activate"
Dim lngTempValue As Long
    'Error Trap
    On Error GoTo Form_Activate_ErrorHandler
    'This section of code causes window to be 'always on top' _
        until it is 'sunk' with another call to Floatwindow, see form_unload event
    If conFloating = True Then
        Me.Show
        lngTempValue = Screen.ActiveForm.hWnd
        Call FloatWindow(lngTempValue, conFloat)
    End If
    'CODE_HERE
Exit Sub
'Error Handler
Form_Activate_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub
Private Sub Form_Load()
'Created by Robin G Brown, INSERT_DATE
'DEFAULT_BEHAVIOUR
'Sub Declares
Const conSub = "Form_Load"
    'Error Trap
    On Error GoTo Form_Load_ErrorHandler
    'CODE_HERE
Exit Sub
'Error Handler
Form_Load_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub
Private Sub Form_Resize()
'Created by Robin G Brown, INSERT_DATE
'DEFAULT_BEHAVIOUR
'Sub Declares
Const conSub = "Form_Resize"
    'Error Trap
    On Error GoTo Form_Resize_ErrorHandler
    'Boilerplate code to center form on screen
    If WindowState > 0 Or conCentering <> True Then
        Exit Sub
    Else
        Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    End If
    'CODE_HERE
Exit Sub
'Error Handler
Form_Resize_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub
Private Sub Form_Unload(Cancel As Integer)
'Created by Robin G Brown, INSERT_DATE
'DEFAULT_BEHAVIOUR
'Function Declares
Const conFunction = "Form_Unload"
Dim lngTempValue As Long
    'Errors 'disabled' to ensure that form unloads
    On Error Resume Next
    'this section of code should also be used whenever the form is hidden
    If conFloating <> False Then
        lngTempValue = Screen.ActiveForm.hWnd
        Call FloatWindow(lngTempValue, conSink)
    End If
    'CODE_HERE
End Sub
