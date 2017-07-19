VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmCSVTester 
   Caption         =   "Form1"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   9900
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTable 
      Height          =   375
      Left            =   60
      TabIndex        =   3
      Top             =   540
      Width           =   3615
   End
   Begin MSFlexGridLib.MSFlexGrid grdDisplay 
      Height          =   6435
      Left            =   60
      TabIndex        =   2
      Top             =   1020
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   11351
      _Version        =   65541
   End
   Begin VB.TextBox txtDirectory 
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   3615
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   855
      Left            =   3780
      TabIndex        =   0
      Top             =   60
      Width           =   6015
   End
End
Attribute VB_Name = "frmCSVTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cdbCSVBase As CSVBase

Private Sub Form_Load()

    Set cdbCSVBase = New CSVBase 'CreateObject("CSVDatabase.CSVBase")
    
End Sub

Private Sub cmdOpen_Click()

Dim ctbTable As CSVTable

    cdbCSVBase.directory = txtDirectory.Text
    Call cdbCSVBase.AddTable(txtTable.Text)
    Set ctbTable = cdbCSVBase.ctbTables(1)
    Call ctbTable.DisplayInGrid(grdDisplay)
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Set cdbCSVBase = Nothing
    
End Sub
