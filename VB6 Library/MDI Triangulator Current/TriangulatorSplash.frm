VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Tri-Systems"
   ClientHeight    =   5805
   ClientLeft      =   1110
   ClientTop       =   1230
   ClientWidth     =   6885
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
   ScaleHeight     =   5805
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraSplash 
      Height          =   5595
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6615
      Begin VB.PictureBox picApplication 
         AutoSize        =   -1  'True
         Height          =   3810
         Left            =   240
         Picture         =   "TriangulatorSplash.frx":0000
         ScaleHeight     =   3750
         ScaleWidth      =   6000
         TabIndex        =   1
         Top             =   300
         Width           =   6060
      End
      Begin VB.Label lblApplication 
         Caption         =   "Tri-Systems"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   120
         TabIndex        =   2
         Top             =   4260
         Width           =   6315
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
'-----------------------------------------------------------------------------
Option Explicit
Private strPictureFile As String

Public Property Let PictureFile(ByVal strNewPictureFile As String)
'Created by Robin G Brown
    'Error trap
    On Error Resume Next
    strPictureFile = strNewPictureFile
End Property

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
    'Set the picture at runtime
    If strPictureFile <> "" Then
        picApplication.Picture = LoadPicture(strPictureFile)
    End If
    Me.Caption = App.CompanyName
    lblApplication = App.CompanyName & " " & App.ProductName
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

