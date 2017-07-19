VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Test"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   495
      Left            =   60
      TabIndex        =   1
      Top             =   600
      Width           =   4575
   End
   Begin VB.TextBox txtTest 
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4575
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lngReturn As Long
Private Thing As Alarm

Private Sub cmdTest_Click()
'Test action

    Set Thing = CreateObject("AlarmTimer.Alarm")
    MsgBox "Pre thing"
    Call Thing.SetAlarm(Me, 1)
    MsgBox "post thing"
    
End Sub

Public Sub AlarmRing()
'Created by Robin G Brown, INSERT_DATE
'This is the return event from the Alarm object
'Sub Declares
    'Error Trap
    On Error Resume Next
    MsgBox "AlarmRing event fired"
End Sub

Private Sub Form_Load()
    Set Thing = Nothing
End Sub
