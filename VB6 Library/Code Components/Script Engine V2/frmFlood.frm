VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Begin VB.Form frmAttention 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Please wait..."
   ClientHeight    =   1275
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
   ScaleHeight     =   1275
   ScaleWidth      =   4605
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
Attribute VB_Name = "frmAttention"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Robin G Brown, 22nd july 1997
'-----------------------------------------------------------------------------
'   This form contains code for _
    diverting the users attention while slow things happen
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "frmAttention"
Private intCounter As Integer
Private intReturn As Integer

Private Sub Form_Activate()
'Created by Robin G Brown
Dim lngTempValue As Long
    On Error Resume Next
    Me.Show
    lngTempValue = Screen.ActiveForm.hWnd
    Call FloatWindow(lngTempValue, 1)
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
    Exit Sub
End Sub

Private Sub Form_Resize()
'Created by Robin G Brown
    On Error Resume Next
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
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
    Call FloatWindow(lngTempValue, 0)
    'Stop the animation control
    aniAttention.AutoPlay = False
    aniAttention.Close
End Sub

