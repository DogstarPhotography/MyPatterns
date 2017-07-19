VERSION 5.00
Begin VB.Form frmBoilerDialog 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Boiler Dialog"
   ClientHeight    =   4455
   ClientLeft      =   1140
   ClientTop       =   1545
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4455
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmBoilerDialog"
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
'   5. Replace all frmBoilerDialog with name of form
'   6. Fill in all DEFAULT_BEHAVIOUR
'   7. Replace strProperty with new property
'   8. Complete DelayedErrorPropertyGet and DelayedErrorPropertyLet
'   9. Delete this SPECIAL_INSTRUCTIONS set of comments
'   10. Save File As... with suitable filename
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "frmBoilerDialog"
Private lngCounter As Long
Private lngReturn As Long
Private strProperty As String

Private Sub Form_Load()
'Created by Robin G Brown, INSERT_DATE
'DEFAULT_BEHAVIOUR
'Sub Declares
Const conSub = "Form_Load"
    'Error Trap
    On Error GoTo Form_Load_ErrorHandler
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    strProperty = ""
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

Property Get DelayedErrorPropertyGet() As String
'Created by Robin G Brown, INSERT_DATE
'DESCRIPTION_OF_FUNCTIONALITY_HERE
'Function Declares
    'Error Trap
    On Error Resume Next
    'CODE_HERE
    DelayedErrorPropertyGet = strProperty
    If Err.Number <> 0 Then
        'ERROR_HANDLING_CODE_HERE
        Err.Clear
    End If
End Property

Property Let DelayedErrorPropertyLet(ByVal strNewProperty As String)
'Created by Robin G Brown, INSERT_DATE
'DESCRIPTION_OF_FUNCTIONALITY_HERE
'Sub Declares
    'Error Trap
    On Error Resume Next
    'CODE_HERE
    strProperty = strNewProperty
    If Err.Number <> 0 Then
        'ERROR_HANDLING_CODE_HERE
        Err.Clear
    End If
End Property

Private Sub Form_Unload(Cancel As Integer)
'Created by Robin G Brown, INSERT_DATE
'DEFAULT_BEHAVIOUR
'Function Declares
Const conFunction = "Form_Unload"
    'Errors 'disabled' to ensure that form unloads
    On Error Resume Next
    'CODE_HERE
    If Err.Number <> 0 Then
        'ERROR_HANDLING_CODE_HERE
        Err.Clear
    End If
End Sub
