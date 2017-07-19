VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmSelectDatabase 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Database"
   ClientHeight    =   1605
   ClientLeft      =   1140
   ClientTop       =   1545
   ClientWidth     =   5925
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
   ScaleHeight     =   1605
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog cdlSelectDatabase 
      Left            =   1860
      Top             =   1020
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.Frame fraDateTime 
      Height          =   1515
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Select"
         Default         =   -1  'True
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   1020
         Width           =   1275
      End
      Begin VB.TextBox txtDatabaseName 
         Height          =   345
         Left            =   120
         TabIndex        =   2
         Top             =   540
         Width           =   5475
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "&OK"
         Height          =   315
         Index           =   0
         Left            =   2940
         TabIndex        =   4
         Top             =   1020
         Width           =   1275
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "&Cancel"
         Height          =   315
         Index           =   1
         Left            =   4320
         TabIndex        =   5
         Top             =   1020
         Width           =   1275
      End
      Begin VB.Label lblSelect 
         Caption         =   "Select &Database"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmSelectDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Robin G Brown, 31st July 1997
'-----------------------------------------------------------------------------
'   This form contains code for _
    a simple dialog for setting a database name
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "frmSelectDatabase"
Private strDatabaseNameProperty As String

Private Sub cmdSelect_Click()
'Created by Robin G Brown, 31/7/97
'Select an Access database
'Sub Declares
Const conSub = "cmdSelect_Click"
    'Error Trap
    On Error GoTo cmdSelect_Click_ErrorHandler
    With cdlSelectDatabase
        .CancelError = True
        .filename = strDatabaseNameProperty
        .Filter = conAccessFilter
        .ShowOpen
        txtDatabaseName.Text = .filename
    End With
Exit Sub
'Error Handler
cmdSelect_Click_ErrorHandler:
    Select Case Err.Number
    Case cdlCancel
        'Cancel pressed, ignore
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Private Sub Form_Load()
'Created by Robin G Brown, 31/7/97
'Sub Declares
    'Error Trap
    On Error Resume Next
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

Private Sub cmdAction_Click(Index As Integer)
'Created by Robin G Brown, 31/7/97
'Handle user actions
'Sub Declares
Const conSub = "cmdAction_Click"
    'Error Trap
    On Error GoTo cmdAction_Click_ErrorHandler
    Select Case Index
    Case 0
        'OK
        strDatabaseNameProperty = txtDatabaseName.Text
        Me.Hide
    Case 1
        'Cancel
        strDatabaseNameProperty = ""
        Me.Hide
    End Select
Exit Sub
'Error Handler
cmdAction_Click_ErrorHandler:
    Select Case Err.Number
    Case Else
        'MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Me.Hide
End Sub

Property Let DatabaseName(ByVal DefaultDatabaseName As String)
'Created by Robin G Brown, 29/7/97
'Return the dialog setting for the date/time
'Function Declares
    'Error Trap
    On Error Resume Next
    strDatabaseNameProperty = DefaultDatabaseName
End Property

Property Get DatabaseName() As String
'Created by Robin G Brown, 29/7/97
'Return the dialog setting for the date/time
'Function Declares
    'Error Trap
    On Error Resume Next
    DatabaseName = strDatabaseNameProperty
End Property

