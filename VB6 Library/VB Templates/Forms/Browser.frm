VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "comctl32.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.0#0"; "shdocvw.dll"
Begin VB.Form frmBrowser 
   ClientHeight    =   5130
   ClientLeft      =   3060
   ClientTop       =   3345
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5130
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   3734
      Left            =   50
      TabIndex        =   0
      Top             =   1215
      Width           =   5393
      Object.Height          =   249
      Object.Width           =   360
      AutoSize        =   0
      ViewMode        =   1
      AutoSizePercentage=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   6180
      Top             =   1500
   End
   Begin ComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   741
      Appearance      =   1
      ImageList       =   "imlIcons"
      BeginProperty Buttons {7791BA41-E020-11CF-8E74-00A0C90F26F8} 
         NumButtons      =   6
         BeginProperty Button1 {7791BA43-E020-11CF-8E74-00A0C90F26F8} 
            Key             =   "Back"
            Object.ToolTipText     =   "Back"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {7791BA43-E020-11CF-8E74-00A0C90F26F8} 
            Key             =   "Forward"
            Object.ToolTipText     =   "Forward"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {7791BA43-E020-11CF-8E74-00A0C90F26F8} 
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {7791BA43-E020-11CF-8E74-00A0C90F26F8} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {7791BA43-E020-11CF-8E74-00A0C90F26F8} 
            Key             =   "Home"
            Object.ToolTipText     =   "Home"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {7791BA43-E020-11CF-8E74-00A0C90F26F8} 
            Key             =   "Search"
            Object.ToolTipText     =   "Search"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
      EndProperty
      MouseIcon       =   "Browser.frx":0000
   End
   Begin VB.PictureBox picAddress 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   6540
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   420
      Width           =   6540
      Begin VB.ComboBox cboAddress 
         Height          =   315
         Left            =   45
         TabIndex        =   2
         Text            =   "¯¯END!"
         Top             =   300
         Width           =   3795
      End
      Begin VB.Label lblAddress 
         Caption         =   "&Address:"
         Height          =   255
         Left            =   45
         TabIndex        =   1
         Tag             =   "&Address:"
         Top             =   60
         Width           =   3075
      End
   End
   Begin ComctlLib.ImageList imlIcons 
      Left            =   2670
      Top             =   2325
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      BeginProperty Images {8556BCD1-E01E-11CF-8E74-00A0C90F26F8} 
         NumListImages   =   6
         BeginProperty ListImage1 {8556BCD3-E01E-11CF-8E74-00A0C90F26F8} 
            Picture         =   "Browser.frx":001C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {8556BCD3-E01E-11CF-8E74-00A0C90F26F8} 
            Picture         =   "Browser.frx":02FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {8556BCD3-E01E-11CF-8E74-00A0C90F26F8} 
            Picture         =   "Browser.frx":05E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {8556BCD3-E01E-11CF-8E74-00A0C90F26F8} 
            Picture         =   "Browser.frx":08C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {8556BCD3-E01E-11CF-8E74-00A0C90F26F8} 
            Picture         =   "Browser.frx":0BA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {8556BCD3-E01E-11CF-8E74-00A0C90F26F8} 
            Picture         =   "Browser.frx":0E86
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public StartingAddress As String
Dim mbDontNavigateNow As Boolean
Private Sub Form_Load()
    On Error Resume Next
    Me.Show
    tbToolBar.Refresh
    Form_Resize

    cboAddress.Move 50, lblAddress.Top + lblAddress.Height + 15

    If Len(StartingAddress) > 0 Then
        cboAddress.Text = StartingAddress
        cboAddress.AddItem cboAddress.Text
        'try to navigate to the starting address
        timTimer.Enabled = True
        brwWebBrowser.Navigate StartingAddress
    End If

End Sub



Private Sub brwWebBrowser_DownloadComplete()
    On Error Resume Next
    Me.Caption = brwWebBrowser.LocationName
End Sub

Private Sub brwWebBrowser_NavigateComplete(ByVal URL As String)
    Dim i As Integer
    Dim bFound As Boolean
    Me.Caption = brwWebBrowser.LocationName
    For i = 0 To cboAddress.ListCount - 1
        If cboAddress.List(i) = brwWebBrowser.LocationURL Then
            bFound = True
            Exit For
        End If
    Next i
    mbDontNavigateNow = True
    If bFound Then
        cboAddress.RemoveItem i
    End If
    cboAddress.AddItem brwWebBrowser.LocationURL, 0
    cboAddress.ListIndex = 0
    mbDontNavigateNow = False
End Sub

Private Sub cboAddress_Click()
    If mbDontNavigateNow Then Exit Sub
    timTimer.Enabled = True
    brwWebBrowser.Navigate cboAddress.Text
End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        cboAddress_Click
    End If
End Sub

Private Sub Form_Resize()
    cboAddress.Width = Me.ScaleWidth - 100
    brwWebBrowser.Width = Me.ScaleWidth - 100
    brwWebBrowser.Height = Me.ScaleHeight - (picAddress.Top + picAddress.Height) - 100
End Sub

Private Sub timTimer_Timer()
    If brwWebBrowser.Busy = False Then
        timTimer.Enabled = False
        Me.Caption = brwWebBrowser.LocationName
    Else
        Me.Caption = "Working..."
    End If
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As Button)
    On Error Resume Next
     
    timTimer.Enabled = True
     
    Select Case Button.Key
        Case "Back"
            brwWebBrowser.GoBack
        Case "Forward"
            brwWebBrowser.GoForward
        Case "Refresh"
            brwWebBrowser.Refresh
        Case "Home"
            brwWebBrowser.GoHome
        Case "Search"
            brwWebBrowser.GoSearch
        Case "Stop"
            timTimer.Enabled = False
            brwWebBrowser.Stop
            Me.Caption = brwWebBrowser.LocationName
    End Select

End Sub

