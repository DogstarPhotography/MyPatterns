VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "comdlg32.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmScriptConfigure 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Script Builder"
   ClientHeight    =   8280
   ClientLeft      =   1140
   ClientTop       =   1545
   ClientWidth     =   12300
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
   ScaleHeight     =   8280
   ScaleWidth      =   12300
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraActions 
      Caption         =   "View Actions"
      Height          =   4335
      Left            =   60
      TabIndex        =   21
      Top             =   3840
      Width           =   12135
      Begin VB.CommandButton cmdAction 
         Caption         =   "24"
         Height          =   375
         Index           =   2
         Left            =   9840
         TabIndex        =   23
         Top             =   3840
         Width           =   2175
      End
      Begin MSDBGrid.DBGrid grdActions 
         Bindings        =   "frmScriptConfigure.frx":0000
         Height          =   3435
         Left            =   120
         OleObjectBlob   =   "frmScriptConfigure.frx":0015
         TabIndex        =   22
         Top             =   300
         Width           =   11895
      End
      Begin VB.Data datActions 
         Caption         =   "Actions"
         Connect         =   "Access"
         DatabaseName    =   "23"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "ActionsParametersValues"
         Top             =   3840
         Width           =   2895
      End
   End
   Begin MSComDlg.CommonDialog cdlScript 
      Left            =   10980
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin VB.Frame fraScriptActions 
      Caption         =   "New Action"
      Height          =   3675
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   12135
      Begin VB.CommandButton cmdAction 
         Caption         =   "Add"
         Height          =   375
         Index           =   0
         Left            =   2400
         TabIndex        =   20
         Top             =   3120
         Width           =   2175
      End
      Begin VB.ComboBox cboParameter 
         Height          =   315
         Index           =   4
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2700
         Width           =   2175
      End
      Begin VB.ComboBox cboParameter 
         Height          =   315
         Index           =   3
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2280
         Width           =   2175
      End
      Begin VB.ComboBox cboParameter 
         Height          =   315
         Index           =   2
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1860
         Width           =   2175
      End
      Begin VB.ComboBox cboParameter 
         Height          =   315
         Index           =   1
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1440
         Width           =   2175
      End
      Begin VB.ComboBox cboCommand 
         Height          =   315
         Index           =   4
         Left            =   10500
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   2700
         Width           =   1515
      End
      Begin VB.ComboBox cboCommand 
         Height          =   315
         Index           =   3
         Left            =   10500
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   2280
         Width           =   1515
      End
      Begin VB.ComboBox cboCommand 
         Height          =   315
         Index           =   2
         Left            =   10500
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1860
         Width           =   1515
      End
      Begin VB.ComboBox cboCommand 
         Height          =   315
         Index           =   1
         Left            =   10500
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1440
         Width           =   1515
      End
      Begin VB.TextBox txtValue 
         Height          =   315
         Index           =   4
         Left            =   2400
         TabIndex        =   18
         Top             =   2700
         Width           =   7995
      End
      Begin VB.TextBox txtValue 
         Height          =   315
         Index           =   3
         Left            =   2400
         TabIndex        =   15
         Top             =   2280
         Width           =   7995
      End
      Begin VB.TextBox txtValue 
         Height          =   315
         Index           =   2
         Left            =   2400
         TabIndex        =   12
         Top             =   1860
         Width           =   7995
      End
      Begin VB.TextBox txtValue 
         Height          =   315
         Index           =   1
         Left            =   2400
         TabIndex        =   8
         Top             =   1440
         Width           =   7995
      End
      Begin VB.TextBox txtAction 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   660
         Width           =   2175
      End
      Begin VB.TextBox txtValue 
         Height          =   315
         Index           =   0
         Left            =   2400
         TabIndex        =   4
         Top             =   660
         Width           =   7995
      End
      Begin VB.Label lblHeader 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Command"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   10500
         TabIndex        =   9
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label lblHeader 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Default"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   2400
         TabIndex        =   7
         Top             =   1080
         Width           =   7995
      End
      Begin VB.Label lblHeader 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Parameter"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label lblHeader 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Action"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   3
         Top             =   300
         Width           =   7995
      End
      Begin VB.Label lblHeader 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmScriptConfigure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Robin G Brown, 23rd July 1997
'-----------------------------------------------------------------------------
'   This form contains code for _
    configuring scripts using the script engine builder
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "frmScriptConfigure"
Private lngCounter As Long
Private lngReturn As Long
Private bldScriptEngineBuilder As scriptengine.Builder

Private Sub Form_Load()
'Created by Robin G Brown, 28/7/97
'Default behaviour
'Sub Declares
Const conSub = "Form_Load"
Dim strActions() As String
Dim intControlCounter As Integer
Dim intItemcounter As Integer
    'Error Trap
    On Error GoTo Form_Load_ErrorHandler
    Set bldScriptEngineBuilder = CreateObject("ScriptEngine.Builder")
    datActions.DatabaseName = bldScriptEngineBuilder.DatabaseName
    'Fill up all those dropdown lists
    For intControlCounter = 1 To 4
        cboParameter(intControlCounter).AddItem ""
        For intItemcounter = 1 To bldScriptEngineBuilder.GetNumberOfParameters
            cboParameter(intControlCounter).AddItem bldScriptEngineBuilder.GetParameter(intItemcounter)
        Next
        cboCommand(intControlCounter).AddItem ""
        For intItemcounter = 1 To bldScriptEngineBuilder.GetNumberOfCommands
            cboCommand(intControlCounter).AddItem bldScriptEngineBuilder.GetCommand(intItemcounter)
        Next
    Next
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
'Created by Robin G Brown, 28/7/97
'Handle user actions
'Sub Declares
Const conSub = "cmdAction_Click"
Dim lngScriptFileNumber As Long
Dim strScriptLine As String
Dim lngAction As Long
Dim intParamCounter As Integer
    'Error Trap
    On Error GoTo cmdAction_Click_ErrorHandler
    Select Case Index
    Case 0
        'Add a new action
        If txtAction.Text = "" Or txtValue(0).Text = "" Then
            MsgBox "You must have an Action Name and an Action Value to add.", vbOKOnly, "Script Configure"
            Exit Sub
        End If
        lngAction = bldScriptEngineBuilder.AddAction(txtAction.Text, txtValue(0).Text)
        If lngAction <> 0 Then
            'Add the parameters
            For intParamCounter = 1 To 4
                If cboParameter(intParamCounter).List(cboParameter(intParamCounter).ListIndex) = "" Then
                    Exit For
                Else
                    lngReturn = bldScriptEngineBuilder.AddActionParameter( _
                        lngAction, _
                        bldScriptEngineBuilder.GetParameterID(cboParameter(intParamCounter).List(cboParameter(intParamCounter).ListIndex)), _
                        intParamCounter, _
                        txtValue(intParamCounter).Text, _
                        cboCommand(intParamCounter).List(cboCommand(intParamCounter).ListIndex))
                End If
            Next
        End If
    Case 2
        'Close
        Unload Me
    Case Else
        'Error - Ignore
    End Select
Exit Sub
'Error Handler
cmdAction_Click_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Created by Robin G Brown, 28/7/97
'Default behaviour
'Function Declares
Const conFunction = "Form_Unload"
    'Errors 'disabled' to ensure that form unloads
    On Error Resume Next
    Set bldScriptEngineBuilder = Nothing
    'SHOW_CALLING_FORM_IF_REQUIRED
End Sub

