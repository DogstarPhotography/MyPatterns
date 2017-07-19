VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "comdlg32.ocx"
Begin VB.Form frmScriptBuilder 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Script Builder"
   ClientHeight    =   7170
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
   Icon            =   "frmScriptBuilder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7170
   ScaleWidth      =   12300
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdlScript 
      Left            =   10560
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin VB.Frame fraScriptActions 
      Caption         =   "Script Elements"
      Height          =   2775
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   12135
      Begin VB.ComboBox cboFloater 
         Height          =   315
         Left            =   4980
         Style           =   2  'Dropdown List
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2160
         Visible         =   0   'False
         Width           =   7995
      End
      Begin VB.TextBox txtParameter 
         Height          =   315
         Index           =   4
         Left            =   2400
         TabIndex        =   13
         Top             =   1920
         Width           =   7995
      End
      Begin VB.CommandButton cmdSetParameter 
         Height          =   315
         Index           =   4
         Left            =   10500
         TabIndex        =   14
         Top             =   1920
         Width           =   1515
      End
      Begin VB.TextBox txtParameter 
         Height          =   315
         Index           =   3
         Left            =   2400
         TabIndex        =   10
         Top             =   1500
         Width           =   7995
      End
      Begin VB.CommandButton cmdSetParameter 
         Height          =   315
         Index           =   3
         Left            =   10500
         TabIndex        =   11
         Top             =   1500
         Width           =   1515
      End
      Begin VB.TextBox txtParameter 
         Height          =   315
         Index           =   2
         Left            =   2400
         TabIndex        =   7
         Top             =   1080
         Width           =   7995
      End
      Begin VB.CommandButton cmdSetParameter 
         Height          =   315
         Index           =   2
         Left            =   10500
         TabIndex        =   8
         Top             =   1080
         Width           =   1515
      End
      Begin VB.TextBox txtParameter 
         Height          =   315
         Index           =   0
         Left            =   2400
         TabIndex        =   2
         Top             =   240
         Width           =   7995
      End
      Begin VB.CommandButton cmdScript 
         Caption         =   "&Append"
         Height          =   315
         Index           =   1
         Left            =   2400
         TabIndex        =   15
         Top             =   2340
         Width           =   1515
      End
      Begin VB.CommandButton cmdSetParameter 
         Height          =   315
         Index           =   1
         Left            =   10500
         TabIndex        =   5
         Top             =   660
         Width           =   1515
      End
      Begin VB.TextBox txtParameter 
         Height          =   315
         Index           =   1
         Left            =   2400
         TabIndex        =   4
         Top             =   660
         Width           =   7995
      End
      Begin VB.CommandButton cmdScript 
         Caption         =   "&Insert"
         Height          =   315
         Index           =   0
         Left            =   4020
         TabIndex        =   16
         Top             =   2340
         Width           =   1515
      End
      Begin VB.ComboBox cboSelectAction 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblArgument 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label lblArgument 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   1500
         Width           =   2175
      End
      Begin VB.Label lblArgument 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label lblArgument 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   660
         Width           =   2175
      End
   End
   Begin VB.Frame fraScript 
      Caption         =   "Script"
      Height          =   4155
      Left            =   60
      TabIndex        =   17
      Top             =   2940
      Width           =   12135
      Begin VB.CommandButton cmdAction 
         Caption         =   "&Save"
         Height          =   315
         Index           =   3
         Left            =   10500
         TabIndex        =   21
         Top             =   1080
         Width           =   1515
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "&Open"
         Height          =   315
         Index           =   2
         Left            =   10500
         TabIndex        =   20
         Top             =   660
         Width           =   1515
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "&New"
         Height          =   315
         Index           =   1
         Left            =   10500
         TabIndex        =   19
         Top             =   240
         Width           =   1515
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "&Close"
         Height          =   315
         Index           =   0
         Left            =   10500
         TabIndex        =   22
         Top             =   3720
         Width           =   1515
      End
      Begin VB.TextBox txtScript 
         Height          =   3795
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   240
         Width           =   10275
      End
   End
End
Attribute VB_Name = "frmScriptBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Robin G Brown, 23rd July 1997
'-----------------------------------------------------------------------------
'   This form contains code for _
    building scripts using the script engine builder
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "frmScriptBuilder"
Private lngCounter As Long
Private lngReturn As Long
Private intArgumentCounter As Integer
Private bldScriptEngineBuilder As Builder
Private booScriptChanged As Boolean
Private bytCurrentParameter As Byte

Private Sub Form_Load()
'Created by Robin G Brown, 23/7/97
'Default behaviour
'Sub Declares
Const conSub = "Form_Load"
Dim strActions() As String
Dim intControlCounter As Integer
    'Error Trap
    On Error GoTo Form_Load_ErrorHandler
    ReDim strActions(1 To 1)
    Set bldScriptEngineBuilder = CreateObject("ScriptEngineV2.Builder")
    For lngCounter = 1 To bldScriptEngineBuilder.GetNumberOfActions
        cboSelectAction.AddItem bldScriptEngineBuilder.GetAction(lngCounter)
    Next
    For intControlCounter = 1 To 4
        'Hide this set of controls
        lblArgument(intControlCounter).Enabled = False
        txtParameter(intControlCounter).Enabled = False
        cmdSetParameter(intControlCounter).Visible = False
    Next
    'Set any variables to default values
    bytCurrentParameter = 0
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

Private Sub cboSelectAction_Click()
'Created by Robin G Brown, 23/7/97
'Create a set of argument/parameter controls, resize the form and fill in the control values values
'Sub Declares
Const conSub = "cboSelectAction_Click"
Dim lngAction As Long
Dim intControlCounter As Integer
    'Error Trap
    On Error GoTo cboSelectAction_Click_ErrorHandler
    'Sort out the action first
    lngAction = bldScriptEngineBuilder.GetActionNumber(cboSelectAction.List(cboSelectAction.ListIndex))
    txtParameter(0).Text = bldScriptEngineBuilder.GetActionValue(lngAction)
    'Now work out how many parameters there are and set up the controls as required
    For intControlCounter = 1 To 4
        lblArgument(intControlCounter).Caption = bldScriptEngineBuilder.GetActionParameter(lngAction, intControlCounter)
        txtParameter(intControlCounter).Text = bldScriptEngineBuilder.GetActionDefault(lngAction, intControlCounter)
        If lblArgument(intControlCounter).Caption <> "" Then
            lblArgument(intControlCounter).Enabled = True
            txtParameter(intControlCounter).Enabled = True
        Else
            lblArgument(intControlCounter).Enabled = False
            txtParameter(intControlCounter).Enabled = False
        End If
        cmdSetParameter(intControlCounter).Caption = bldScriptEngineBuilder.GetActionCommand(lngAction, intControlCounter)
        If cmdSetParameter(intControlCounter).Caption <> "" Then
            If cmdSetParameter(intControlCounter).Caption = "Server" Then
                cmdSetParameter(intControlCounter).Visible = False
            Else
                cmdSetParameter(intControlCounter).Visible = True
            End If
        Else
            cmdSetParameter(intControlCounter).Visible = False
        End If
    Next
Exit Sub
'Error Handler
cboSelectAction_Click_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Private Sub cmdAction_Click(Index As Integer)
'Created by Robin G Brown, 23/7/97
'Handle user actions
'Sub Declares
Const conSub = "cmdAction_Click"
Dim lngScriptFileNumber As Long
Dim strScriptLine As String
    'Error Trap
    On Error GoTo cmdAction_Click_ErrorHandler
    Select Case Index
    Case 0
        'Close
        Unload Me
    Case 1
        'New
        If booScriptChanged = True Then
            lngReturn = MsgBox("Current script not saved, discard changes?", vbOKCancel, "Script Builder")
            If lngReturn = vbCancel Then
                Exit Sub
            End If
        End If
        txtScript.Text = ""
        booScriptChanged = False
    Case 2
        'Open
        cdlScript.Filter = conAnyTXTFilter
        cdlScript.ShowOpen
        txtScript.Text = ""
        lngScriptFileNumber = FreeFile()
        Open cdlScript.filename For Input As #lngScriptFileNumber
        Do While Not EOF(lngScriptFileNumber)
            Line Input #lngScriptFileNumber, strScriptLine
            txtScript.Text = txtScript.Text & strScriptLine & vbCrLf
        Loop
        Close #lngScriptFileNumber
        booScriptChanged = False
    Case 3
        'Save
        If booScriptChanged = False Then
            lngReturn = MsgBox("Script has not been changed, save anyway?", vbOKCancel, "Script Control")
            If lngReturn = vbCancel Then
                Exit Sub
            End If
        End If
        cdlScript.Filter = conAnyTXTFilter
        cdlScript.ShowSave
        If CheckForDuplicate(cdlScript.filename) = True Then
            'Save the file
            lngScriptFileNumber = FreeFile()
            Open cdlScript.filename For Output As #lngScriptFileNumber
            Print #lngScriptFileNumber, txtScript.Text
            Close #lngScriptFileNumber
        End If
        booScriptChanged = False
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

Private Sub cmdScript_Click(Index As Integer)
'Created by Robin G Brown, 23/7/97
'Add the current script action to the script in the appropriate place
'Sub Declares
Const conSub = "cmdAdd_Click"
Dim strScriptAction As String
Dim strFirstText As String
Dim strLastText As String
    'Error Trap
    On Error GoTo cmdAdd_Click_ErrorHandler
    strScriptAction = txtParameter(0).Text & vbCrLf
    For intArgumentCounter = 1 To 4
        If lblArgument(intArgumentCounter).Caption <> "" Then
            strScriptAction = _
                strScriptAction _
                & lblArgument(intArgumentCounter).Caption _
                & txtParameter(intArgumentCounter).Text _
                & vbCrLf
        End If
    Next
    Select Case Index
    Case 0
        'Insert
        If txtScript.SelStart > 1 Then
            strFirstText = Left$(txtScript.Text, txtScript.SelStart - 1)
        Else
            strFirstText = ""
        End If
        strLastText = Right$(txtScript.Text, Len(txtScript.Text) - txtScript.SelStart)
        txtScript.Text = strFirstText & strScriptAction & strLastText
    Case 1
        'Append
        txtScript.Text = txtScript.Text & vbCrLf & strScriptAction
    End Select
    booScriptChanged = True
Exit Sub
'Error Handler
cmdAdd_Click_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Private Sub cmdSetParameter_Click(Index As Integer)
'Created by Robin G Brown, 23/7/97
'Carry out an action based on the caption of the command button
'Sub Declares
Const conSub = "cmdSetParameter_Click"
Dim intControlCounter As Integer
Dim sdldialog As Dialog
    'Error Trap
    On Error GoTo cmdSetParameter_Click_ErrorHandler
    Select Case cmdSetParameter(Index).Caption
    Case "Program"
        'Load script using common dialog
        cdlScript.Filter = bldScriptEngineBuilder.GetCommandFilter("Program")
        cdlScript.ShowOpen
        txtParameter(Index).Text = cdlScript.filename
    Case "File"
        'Load script using common dialog
        cdlScript.Filter = bldScriptEngineBuilder.GetCommandFilter("File")
        cdlScript.ShowOpen
        txtParameter(Index).Text = cdlScript.filename
    Case "Text"
        'Load script using common dialog
        cdlScript.Filter = bldScriptEngineBuilder.GetCommandFilter("Text")
        cdlScript.ShowOpen
        txtParameter(Index).Text = cdlScript.filename
    Case "CSV"
        'Load script using common dialog
        cdlScript.Filter = bldScriptEngineBuilder.GetCommandFilter("CSV")
        cdlScript.ShowOpen
        txtParameter(Index).Text = cdlScript.filename
    Case "Access"
        'Load script using common dialog
        cdlScript.Filter = bldScriptEngineBuilder.GetCommandFilter("Access")
        cdlScript.ShowOpen
        txtParameter(Index).Text = cdlScript.filename
    Case "Server"
        'Ignore this , it is handled by the GotFocus event of txtParameter
    Case "ODBC"
        'Set an ODBC DSN
        Beep
    Case "Date"
        'Set a date and time
        Set sdldialog = New Dialog
        With sdldialog
            .DateTime = Format$(Now, "Short Date") & " " & Format$(0, "Short Time")
            .ShowDateTime
            txtParameter(Index).Text = Format$(.DateTime, "Short Date") _
                & " " & Format$(.DateTime, "Short Time")
        End With
        Set sdldialog = Nothing
    Case "Time"
        'Set a time
        Set sdldialog = New Dialog
        With sdldialog
            .DateTime = Format$(0, "Short Time")
            .ShowTime
            txtParameter(Index).Text = Format$(.DateTime, "Short Time")
        End With
        Set sdldialog = Nothing
    Case Else
        'Error - ignore
        Beep
    End Select
Exit Sub
'Error Handler
cmdSetParameter_Click_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Private Sub cboFloater_Click()
'Created by Robin G Brown, 29/7/97
'Set the current txtParameter and hide cboFloater
'Sub Declares
    'Error Trap
    On Error Resume Next
    If bytCurrentParameter > 0 Then
        txtParameter(bytCurrentParameter).Text = cboFloater.List(cboFloater.ListIndex)
    End If
    cboFloater.Visible = False
    bytCurrentParameter = 0
End Sub

Private Sub cboFloater_LostFocus()
'Created by Robin G Brown, 29/7/97
'Hide cbofloater
'Sub Declares
    'Error Trap
    On Error Resume Next
    cboFloater.Visible = False
End Sub

Private Sub txtParameter_GotFocus(Index As Integer)
'Created by Robin G Brown, 29/7/97
'In certain cases, cbofloater
'Sub Declares
Dim intCounter As Integer
    'Error Trap
    On Error GoTo txtParameter_GotFocus_ErrorHandler
    If cmdSetParameter(Index).Caption = "Server" And Index > 0 Then
        'Set a Server
        With cboFloater
            .Clear
            .AddItem ""
            For intCounter = 1 To bldScriptEngineBuilder.GetNumberOfServers
                .AddItem bldScriptEngineBuilder.GetServer(intCounter)
                If .List(intCounter) = txtParameter(Index).Text Then
                    .ListIndex = intCounter
                End If
            Next
            .Move txtParameter(Index).Left, txtParameter(Index).Top
            .Visible = True
        End With
        bytCurrentParameter = Index
        cboFloater.SetFocus
    Else
        cboFloater.Visible = False
    End If
Exit Sub
txtParameter_GotFocus_ErrorHandler:
    Select Case Err.Number
    Case 340
        'Control array element 'item' doesn't exist (Error 340) - Do nothing, just ignore
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule
    End Select
    cboFloater.Visible = False
    Exit Sub
End Sub

Private Sub txtParameter_LostFocus(Index As Integer)
'Created by Robin G Brown, 29/7/97
'!DO NOT ADD CODE TO THIS PROCEDURE! - See txtParameter_GotFocus event
End Sub

Private Sub txtScript_Change()
'Created by Robin G Brown, 23/7/97
'Sub Declares
    'Error Trap
    On Error Resume Next
    booScriptChanged = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Created by Robin G Brown, 23/7/97
'Default behaviour
'Function Declares
Const conFunction = "Form_Unload"
    'Errors 'disabled' to ensure that form unloads
    On Error Resume Next
    If booScriptChanged = True Then
        lngReturn = MsgBox("Current script not saved, discard changes?", vbOKCancel, "Script Builder")
        If lngReturn = vbCancel Then
            Cancel = 1
            Exit Sub
        End If
    End If
    Set bldScriptEngineBuilder = Nothing
    'SHOW_CALLING_FORM_IF_REQUIRED
End Sub

