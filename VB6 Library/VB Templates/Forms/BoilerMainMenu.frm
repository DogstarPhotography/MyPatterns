VERSION 5.00
Begin VB.Form frmMainMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Menu"
   ClientHeight    =   3900
   ClientLeft      =   1140
   ClientTop       =   1545
   ClientWidth     =   3000
   ControlBox      =   0   'False
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
   ScaleHeight     =   3900
   ScaleWidth      =   3000
   Begin VB.Frame fraMenu 
      Height          =   3735
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2835
      Begin VB.CommandButton cmdMenu 
         Caption         =   "Menu 3"
         Height          =   555
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   2595
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "Menu 2"
         Height          =   555
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   900
         Width           =   2595
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "Menu 1"
         Height          =   555
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2595
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "E&xit"
         Height          =   555
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   3060
         Width           =   2595
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   2700
         Y1              =   2940
         Y2              =   2940
      End
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Robin G Brown, INSERT_LONG_DATE
'-----------------------------------------------------------------------------
'   This form contains code for _
    HANDLING_USER_MENU_SELECTIONS
'-----------------------------------------------------------------------------
'   SPECIAL_INSTRUCTIONS
'   1. Fill in HANDLING_USER_MENU_SELECTIONS above
'   2. Replace all INSERT_DATE with date of creation
'   3. Fill in all DEFAULT_BEHAVIOUR
'   4. Replace Menu 1, 2 & 3 with suitable functionality
'   5. Delete this SPECIAL_INSTRUCTIONS set of comments
'   6. Save File As... with suitable filename
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "frmMainMenu"
Private intCounter As Integer
Private intReturn As Integer

Private Sub cmdMenu_Click(Index As Integer)
'Created by Robin G Brown, INSERT_DATE
'DEFAULT_BEHAVIOUR
'Sub Declares
    'Error Trap
    On Error Resume Next
    Select Case Index
    Case 0
        'Exit
        Unload Me
    Case 1
        'Menu 1
        'CODE_HERE
    Case 2
        'Menu 2
        'CODE_HERE
    Case 3
        'Menu 3
        'CODE_HERE
    Case Else
        'Error - do nothing
    End Select
End Sub

Private Sub Form_Activate()
'Created by Robin G Brown, INSERT_DATE
'DEFAULT_BEHAVIOUR
'Sub Declares
Const conSub = "Form_Activate"
Dim lngTempValue As Long
    'Error Trap
    On Error GoTo Form_Activate_ErrorHandler
    'CODE_HERE
Exit Sub
'Error Handler
Form_Activate_ErrorHandler:
    Select Case Err
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Error$(Err), vbExclamation, conModule & " : " & conSub
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
    Select Case Err
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Error$(Err), vbExclamation, conModule & " : " & conSub
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
    'Code to centre screen on window (delete as applicable)
    If Me.WindowState = 0 Then
        Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    End If
    'CODE_HERE
Exit Sub
'Error Handler
Form_Resize_ErrorHandler:
    Select Case Err
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Error$(Err), vbExclamation, conModule & " : " & conSub
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
    'CODE_HERE
End Sub
