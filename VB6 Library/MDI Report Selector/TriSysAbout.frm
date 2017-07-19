VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   5460
   ClientLeft      =   1155
   ClientTop       =   1560
   ClientWidth     =   8610
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
   ScaleHeight     =   5460
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraSplash 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8355
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   3300
         TabIndex        =   5
         Top             =   4740
         Width           =   1755
      End
      Begin VB.PictureBox picApplication 
         Height          =   2895
         Left            =   120
         Picture         =   "TriSysAbout.frx":0000
         ScaleHeight     =   2835
         ScaleWidth      =   7995
         TabIndex        =   1
         Top             =   240
         Width           =   8055
      End
      Begin VB.Label lblCompany 
         Alignment       =   1  'Right Justify
         Caption         =   "Tri-Systems"
         Height          =   555
         Left            =   3840
         TabIndex        =   4
         Top             =   4080
         Width           =   4335
      End
      Begin VB.Label lblVersion 
         Caption         =   "Version Information"
         Height          =   555
         Left            =   120
         TabIndex        =   3
         Top             =   4080
         Width           =   3615
      End
      Begin VB.Label lblApplication 
         Caption         =   "Report Selector"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   3360
         Width           =   8055
      End
   End
End
Attribute VB_Name = "frmAbout"
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
'-----------------------------------------------------------------------------
Option Explicit

Private Sub cmdOK_Click()
'Created by Robin G Brown
    On Error Resume Next
    Unload Me
End Sub

Private Sub Form_Activate()
'Created by Robin G Brown
Dim lngTempValue As Long
    On Error Resume Next
    Me.Show
    lngTempValue = Screen.ActiveForm.hWnd
    Call FloatWindow(lngTempValue, conFloat)
End Sub
Private Sub Form_Resize()
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
