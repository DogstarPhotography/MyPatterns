VERSION 5.00
Begin VB.Form frmDateTime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Date/Time"
   ClientHeight    =   1230
   ClientLeft      =   1140
   ClientTop       =   1545
   ClientWidth     =   4845
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
   ScaleHeight     =   1230
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraDateTime 
      Height          =   1155
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.TextBox txtTime 
         Height          =   315
         Left            =   1260
         MaxLength       =   24
         TabIndex        =   4
         Top             =   660
         Width           =   1875
      End
      Begin VB.TextBox txtDate 
         Height          =   315
         Left            =   1260
         MaxLength       =   24
         TabIndex        =   2
         Top             =   240
         Width           =   1875
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "&Close"
         Height          =   315
         Index           =   1
         Left            =   3240
         TabIndex        =   6
         Top             =   660
         Width           =   1275
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "&OK"
         Height          =   315
         Index           =   0
         Left            =   3240
         TabIndex        =   5
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label lblTime 
         Caption         =   "Enter &Time"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblDate 
         Caption         =   "Enter &Date"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmDateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Robin G Brown, 29th July 1997
'-----------------------------------------------------------------------------
'   This form contains code for _
    simple dialog to set a date and time
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "frmDateTime"
Private lngReturn As Long
Private dteDateTimeValue As Date
Private intShowState As Integer
Private Const conDateAndTime = 0
Private Const conDateOnly = 1
Private Const conTimeOnly = 2

Private Sub Form_Load()
'Created by Robin G Brown, 29/7/97
'Show the form
'Sub Declares
Const conSub = "Form_Load"
    'Error Trap
    On Error GoTo Form_Load_ErrorHandler
    intShowState = conDateAndTime
    txtDate.Text = Format$(Now, "Short Date")
    txtTime.Text = Format$(Now, "Short Time")
    'Me.Show
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
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

Private Sub cmdAction_Click(Index As Integer)
'Created by Robin G Brown, 29/7/97
'Handle user actions
'Sub Declares
Const conSub = "cmdAction_Click"
    'Error Trap
    On Error GoTo cmdAction_Click_ErrorHandler
    Select Case Index
    Case 0
        'OK
        Select Case intShowState
        Case conDateOnly
            dteDateTimeValue = CDate(txtDate.Text)
        Case conTimeOnly
            dteDateTimeValue = CDate(txtTime.Text)
        Case Else
            dteDateTimeValue = CDate(txtDate.Text & " " & txtTime.Text)
        End Select
        Me.Hide
    Case 1
        'Close
        dteDateTimeValue = ""
        Me.Hide
    End Select
Exit Sub
'Error Handler
cmdAction_Click_ErrorHandler:
    Select Case Err.Number
    Case 13
        'Type mismatch
        lngReturn = MsgBox("Incorrect date entered, carry on?", vbOKCancel, "Date/time")
        If lngReturn = vbCancel Then
            Exit Sub
        End If
    Case Else
        'MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Me.Hide
End Sub

Public Sub ShowDateAndTime()
'Created by Robin G Brown, 29/7/97
'Show only the date
'Function Declares
    'Error Trap
    On Error Resume Next
    lblDate.Visible = True
    txtDate.Visible = True
    lblTime.Visible = True
    txtTime.Visible = True
    intShowState = conDateAndTime
End Sub

Public Sub ShowDateOnly()
'Created by Robin G Brown, 29/7/97
'Show only the date
'Function Declares
    'Error Trap
    On Error Resume Next
    lblDate.Visible = True
    txtDate.Visible = True
    lblTime.Visible = False
    txtTime.Visible = False
    intShowState = conDateOnly
End Sub

Public Sub ShowTimeOnly()
'Created by Robin G Brown, 29/7/97
'Show only the date
'Function Declares
    'Error Trap
    On Error Resume Next
    lblDate.Visible = False
    txtDate.Visible = False
    lblTime.Visible = True
    txtTime.Visible = True
    intShowState = conTimeOnly
End Sub

Property Let DateTime(ByVal DefaultDateTime As Date)
'Created by Robin G Brown, 29/7/97
'Return the dialog setting for the date/time
'Function Declares
    'Error Trap
    On Error Resume Next
    dteDateTimeValue = DefaultDateTime
    txtDate.Text = Format$(dteDateTimeValue, "Short Date")
    txtTime.Text = Format$(dteDateTimeValue, "Short Time")
End Property

Property Get DateTime() As Date
'Created by Robin G Brown, 29/7/97
'Return the dialog setting for the date/time
'Function Declares
    'Error Trap
    On Error Resume Next
    DateTime = dteDateTimeValue
End Property

