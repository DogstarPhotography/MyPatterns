VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Begin VB.Form frmFlood 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Please wait..."
   ClientHeight    =   1725
   ClientLeft      =   1140
   ClientTop       =   1545
   ClientWidth     =   4605
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
   ScaleHeight     =   1725
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.ProgressBar pgbFlood 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      _Version        =   327680
      Appearance      =   1
   End
   Begin ComCtl2.Animation aniAttention 
      Height          =   1035
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1826
      _Version        =   327680
      AutoPlay        =   -1  'True
      BackColor       =   12632256
      FullWidth       =   289
      FullHeight      =   69
   End
End
Attribute VB_Name = "frmFlood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Robin G Brown, 30th April 1997
'-----------------------------------------------------------------------------
'   This form contains code for _
    diverting the users attention while slow things happen
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "frmFlood"

Private Sub Form_Activate()
'Created by Robin G Brown
Dim lngTempValue As Long
    On Error Resume Next
    Me.Show
    lngTempValue = Screen.ActiveForm.hWnd
    Call FloatWindow(lngTempValue, conFloat)
    'Play the animation
    aniAttention.AutoPlay = True
End Sub

Private Sub Form_Load()
'Created by Robin G Brown, 30/4/97
'Default Behaviour
'Sub Declares
Const conSub = "Form_Load"
    'Error Trap
    On Error GoTo Form_Load_ErrorHandler
    'Load the animation control with an .avi file
    aniAttention.Open App.Path & "\filemove.avi"
    aniAttention.AutoPlay = True
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

Private Sub Form_Unload(Cancel As Integer)
'Created by Robin G Brown, 30/4/97
'Default Behaviour
'Function Declares
Const conFunction = "Form_Unload"
Dim lngTempValue As Long
    'Errors 'disabled' to ensure that form unloads
    On Error Resume Next
    lngTempValue = Screen.ActiveForm.hWnd
    Call FloatWindow(lngTempValue, conSink)
    'Stop the animation control
    aniAttention.AutoPlay = False
    aniAttention.Close
End Sub
Public Sub SetFlood(ByVal intPercentage As Integer)
'Created by Robin G Brown, 30/4/97
'Set the amount of floodbar to be shown
'Sub Declares
    'Error Trap
    On Error Resume Next
    pgbFlood.Value = intPercentage
    DoEvents
End Sub

