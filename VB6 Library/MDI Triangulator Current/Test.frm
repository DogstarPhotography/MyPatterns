VERSION 5.00
Begin VB.Form FrmTest 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDestroy 
      Caption         =   "Destroy"
      Height          =   495
      Left            =   1020
      TabIndex        =   1
      Top             =   1380
      Width           =   1935
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create"
      Height          =   495
      Left            =   1020
      TabIndex        =   0
      Top             =   660
      Width           =   1935
   End
End
Attribute VB_Name = "FrmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim trpReporter As Object

Private Sub cmdCreate_Click()

    Set trpReporter = CreateObject("TriReporter.Application")
    trpReporter.Visible (True)
    
End Sub

Private Sub cmdDestroy_Click()

    'Set trpReporter = Nothing
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
        Set trpReporter = Nothing

End Sub
