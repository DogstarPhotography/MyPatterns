VERSION 5.00
Begin VB.Form frmTestMessage 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ScriptEngine V2"
   ClientHeight    =   1080
   ClientLeft      =   1140
   ClientTop       =   1545
   ClientWidth     =   5250
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1080
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Close"
      Height          =   435
      Left            =   1980
      TabIndex        =   1
      Top             =   540
      Width           =   1215
   End
   Begin VB.Label lblMessage 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5115
   End
End
Attribute VB_Name = "frmTestMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Robin G Brown, 10th September 1997
'-----------------------------------------------------------------------------
'   This form contains code for _
    a simple test message
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "frmTestMessage"
'Declare constants for boilerplate code
Private Const conCentering = True    'causes form to center on screen when shown, check form_resize event
Private Const conFloating = True    'causes form to be 'always on top', check form_activate and form_unload events for other bits

Private Sub cmdOK_Click()
'Created by Robin G Brown, 10/9/97
    'Error Trap
    On Error Resume Next
    Unload Me
End Sub

Private Sub Form_Activate()
'Created by Robin G Brown, 10/9/97
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
Exit Sub
'Error Handler
Form_Activate_ErrorHandler:
    MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    Exit Sub
End Sub

Private Sub Form_Resize()
'Created by Robin G Brown, 10/9/97
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
Exit Sub
'Error Handler
Form_Resize_ErrorHandler:
    MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Created by Robin G Brown, 10/9/97
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
End Sub
