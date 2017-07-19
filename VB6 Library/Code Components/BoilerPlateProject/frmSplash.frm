VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Tri-Systems"
   ClientHeight    =   4335
   ClientLeft      =   1110
   ClientTop       =   1230
   ClientWidth     =   7230
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
   ScaleHeight     =   4335
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraSplash 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin VB.PictureBox picApplication 
         Height          =   2895
         Left            =   120
         ScaleHeight     =   2835
         ScaleWidth      =   6675
         TabIndex        =   1
         Top             =   240
         Width           =   6735
      End
      Begin VB.Label lblApplication 
         Caption         =   "Application Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   2
         Top             =   3360
         Width           =   6735
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Robin G Brown
'-----------------------------------------------------------------------------
'   This form contains code for _
    a simple 'Splash' form, it requires no code changes, _
    merely the addition of a nice picture and some appropriate text
'   Note that the splash form will always appear centered on the screen _
    and 'always on top'
'   How about an animated splashscreen? - RGB/14/8/97
'-----------------------------------------------------------------------------
Option Explicit

Private Sub Form_Activate()
'Created by Robin G Brown
Dim lngTempValue As Long
    On Error Resume Next
    Me.Show
    lngTempValue = Screen.ActiveForm.hWnd
    Call FloatWindow(lngTempValue, conFloat)
End Sub

Private Sub Form_Load()
'Created by Robin G Brown
    On Error Resume Next
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Created by Robin G Brown
Dim lngTempValue As Long
    'Errors 'disabled' to ensure that form unloads
    On Error Resume Next
    lngTempValue = Screen.ActiveForm.hWnd
    Call FloatWindow(lngTempValue, conSink)
End Sub
