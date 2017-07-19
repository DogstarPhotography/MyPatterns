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
'   7. Delete this SPECIAL_INSTRUCTIONS set of comments
'   8. Save File As... with suitable filename
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "frmBoilerPlate"
Private lngCounter As Long
Private lngReturn As Long

Private Sub Form_Initialize()
'Created by Robin G Brown, INSERT_DATE
'DESCRIPTION_OF_SUBROUTINE_FUNCTIONALITY_HERE
'Sub Declares
    'Error Trap
    On Error Resume Next
    'CODE_HERE
    strPROPERTY = ""
    Set PROPERTYOBJECT = Nothing
    If Err.Number <> 0 Then
        'ERROR_HANDLING_CODE_HERE
        Err.Raise Err.Number, conModule & ":Initialize", "Unexpected error : " & Err.Description
        Err.Clear
    End If
End Sub

'-----------------------------------------------------------------------------
'   Form events
'-----------------------------------------------------------------------------

Private Sub Form_Activate()
'Created by Robin G Brown, INSERT_DATE
'DEFAULT_BEHAVIOUR
'Sub Declares
Const conSub = "Form_Activate"
    'Error Trap
    On Error GoTo Form_Activate_ErrorHandler
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
    'Code to centre screen on window (delete as applicable)
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
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
    'Resize controls, unless form is minimized
    If Me.WindowState <> 1 Then
        'Control.Height = Me.Height - OFFSET
        'Control.Width = Me.Width - OFFSET
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
Const conSub = "Form_Unload"
    'Errors 'disabled' to ensure that form unloads
    On Error Resume Next
    'CODE_HERE
    If Err.Number <> 0 Then
        'ERROR_HANDLING_CODE_HERE
        Err.Clear
    End If
End Sub

'-----------------------------------------------------------------------------
'   Functions and Methods
'-----------------------------------------------------------------------------

'-----------------------------------------------------------------------------
'   Properties
'-----------------------------------------------------------------------------

Private Sub Form_Terminate()
'Created by Robin G Brown, INSERT_DATE
'DESCRIPTION_OF_SUBROUTINE_FUNCTIONALITY_HERE
'Sub Declares
    'Error Trap
    On Error Resume Next
    Err.Clear
    'CODE_HERE
    Set PROPERTYOBJECT = Nothing
    If Err.Number <> 0 Then
        'ERROR_HANDLING_CODE_HERE
    End If
End Sub



