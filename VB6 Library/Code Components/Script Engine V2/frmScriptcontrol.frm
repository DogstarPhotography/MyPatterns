VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "comdlg32.ocx"
Begin VB.Form frmScriptControl 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Script Control"
   ClientHeight    =   5880
   ClientLeft      =   2265
   ClientTop       =   2925
   ClientWidth     =   10320
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
   Icon            =   "frmScriptcontrol.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5880
   ScaleWidth      =   10320
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAction 
      Caption         =   "Actions"
      Height          =   3255
      Left            =   8040
      TabIndex        =   5
      Top             =   60
      Width           =   2175
      Begin VB.CommandButton cmdScriptAction 
         Caption         =   "&Edit"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CommandButton cmdScriptAction 
         Caption         =   "&Open"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   660
         Width           =   1935
      End
      Begin VB.CommandButton cmdScriptAction 
         Caption         =   "&Close"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   2760
         Width           =   1935
      End
      Begin VB.CommandButton cmdScriptAction 
         Caption         =   "&Save"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   1500
         Width           =   1935
      End
      Begin VB.CommandButton cmdScriptAction 
         Caption         =   "&Run"
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   10
         Top             =   1920
         Width           =   1935
      End
      Begin VB.CommandButton cmdScriptAction 
         Caption         =   "&Load and Run"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame fraScript 
      Caption         =   "Script"
      Height          =   5715
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   7875
      Begin VB.TextBox txtScript 
         Height          =   4755
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   7635
      End
      Begin VB.TextBox txtScriptLocation 
         Height          =   315
         Left            =   1380
         TabIndex        =   2
         Top             =   240
         Width           =   6375
      End
      Begin MSComDlg.CommonDialog cdlScript 
         Left            =   7080
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   327681
      End
      Begin VB.Label lblScript 
         Caption         =   "File &Contents"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2475
      End
      Begin VB.Label lblScript 
         Caption         =   "&File Name"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmScriptControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Robin G Brown, 7th July 1997
'-----------------------------------------------------------------------------
'   This form contains code for _
    controlling scripts and the Script Engine
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "frmScriptControl"
Private lngReturn As Long
Private booScriptChanged As Boolean
Private WithEvents sceScriptEngine As Engine
Attribute sceScriptEngine.VB_VarHelpID = -1

Private Sub sceScriptEngine_ScriptCompleted()
'Created by Robin G Brown, 9/9/97
'Handle the ScriptCompleted event of the ScriptEngine Object
'Sub Declares
    'Error Trap
    On Error Resume Next
    'Make some noise
    Beep
    Beep
    Beep
End Sub

Private Sub sceScriptEngine_ScriptUpdate(ByVal strMessage As String, ByVal intPercentDone As Integer)
'Created by Robin G Brown, 9/9/97
'Handle the ScriptUpdate event of the ScriptEngine Object
'Sub Declares
    'Error Trap
    On Error Resume Next
    'Ignore updates
End Sub

Private Sub cmdScriptAction_Click(Index As Integer)
'Created by Robin G Brown, 9/7/97
'Handle user actions
'Sub Declares
Const conSub = "cmdScriptAction_Click"
Dim lngScriptFileNumber As Long
Dim strScriptLine As String
Dim strTempScript As String
    'Error Trap
    On Error GoTo cmdScriptAction_Click_ErrorHandler
    cdlScript.CancelError = True
    Select Case Index
    Case 0
        'Close, check for changes to the script
        If booScriptChanged = True Then
            lngReturn = MsgBox("This script has changed, save changes?", vbYesNoCancel, "Script Control")
            If lngReturn = vbCancel Then
                Exit Sub
            ElseIf lngReturn = vbNo Then
                'Do not save
            Else
                'Save
                Call SaveScriptFile
                txtScript.Enabled = False
                booScriptChanged = False
            End If
        End If
        Unload Me
    Case 1
        'Load script using common dialog
        cdlScript.Filter = conAnyTXTFilter
        cdlScript.ShowOpen
        txtScriptLocation.Text = cdlScript.filename
        Call LoadScriptFile
        Screen.MousePointer = vbHourglass
        'Process script
        Call sceScriptEngine.RunText(txtScript.Text)
        txtScript.Text = ""
        txtScript.Enabled = False
        booScriptChanged = False
        Screen.MousePointer = vbDefault
        Beep
    Case 2
        'View - open the script file and show it in the text box
        cdlScript.Filter = conAnyTXTFilter
        cdlScript.ShowOpen
        txtScriptLocation.Text = cdlScript.filename
        Call LoadScriptFile
        txtScript.Enabled = False
        booScriptChanged = False
    Case 3
        'Edit - enable the text box for changes
        txtScript.Enabled = True
        booScriptChanged = False
    Case 4
        'Save changes
        If booScriptChanged = False Then
            lngReturn = MsgBox("Script has not been changed, save anyway?", vbOKCancel, "Script Control")
            If lngReturn <> vbOK Then
                Exit Sub
            End If
        End If
        cdlScript.Filter = conAnyTXTFilter
        cdlScript.ShowSave
        txtScriptLocation.Text = cdlScript.filename
        If CheckForDuplicate(txtScriptLocation.Text) = True Then
            'Save
            Call SaveScriptFile
            txtScript.Enabled = False
            booScriptChanged = False
        End If
    Case 5
        'Run
        If booScriptChanged = True Then
            lngReturn = MsgBox("This script has changed, only saved changes will be processed, save changes?", vbYesNoCancel, "Script Control")
            If lngReturn = vbCancel Then
                Exit Sub
            ElseIf lngReturn = vbNo Then
                'Do not save
            Else
                'Save
                Call SaveScriptFile
                txtScript.Enabled = False
                booScriptChanged = False
            End If
        End If
        'Process script
        Screen.MousePointer = vbHourglass
        Call sceScriptEngine.RunText(txtScript.Text)
        Screen.MousePointer = vbDefault
    Case Else
        'Error - do nothing
    End Select
Exit Sub
'Error Handler
cmdScriptAction_Click_ErrorHandler:
    Select Case Err.Number
    Case cdlCancel
        'Cancel on common Dialog pressed - ignore
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Public Sub LoadScriptFile()
'Created by Robin G Brown, 6/8/97
'Load a script file into the text box
'Sub Declares
Const conSub = "LoadScriptFile"
Dim lngScriptFileNumber As Long
Dim strScriptLine As String
    'Error Trap
    On Error GoTo LoadScriptFile_ErrorHandler
    txtScript.Text = ""
    lngScriptFileNumber = FreeFile()
    Open txtScriptLocation.Text For Input As #lngScriptFileNumber
    Do While Not EOF(lngScriptFileNumber)
        Line Input #lngScriptFileNumber, strScriptLine
        txtScript.Text = txtScript.Text & strScriptLine & vbCrLf
    Loop
    Close #lngScriptFileNumber
Exit Sub
'Error Handler
LoadScriptFile_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Public Sub SaveScriptFile()
'Created by Robin G Brown, 6/8/97
'Save the script file
'Sub Declares
Const conSub = "SaveScriptFile"
Dim lngScriptFileNumber As Long
    'Error Trap
    On Error GoTo SaveScriptFile_ErrorHandler
    lngScriptFileNumber = FreeFile()
    Open txtScriptLocation.Text For Output As #lngScriptFileNumber
    Print #lngScriptFileNumber, txtScript.Text
    Close #lngScriptFileNumber
Exit Sub
'Error Handler
SaveScriptFile_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Private Sub Form_Load()
'Created by Robin G Brown, 7/7/97
'Sub Declares
Const conSub = "Form_Load"
    'Error Trap
    On Error GoTo Form_Load_ErrorHandler
    booScriptChanged = False
    'Create a server
    Set sceScriptEngine = New Engine
Exit Sub
'Error Handler
Form_Load_ErrorHandler:
    MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    Exit Sub
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Created by Robin G Brown, 21/7/97
'Check for scripts waiting to be scheduled before closing
'Function Declares
Const conFunction = "Form_Unload"
Dim lngTempValue As Long
    'Errors 'disabled' to ensure that form unloads
    On Error GoTo Form_QueryUnload_ErrorHandler
    If sceScriptEngine.AlarmWaiting = True Then
        lngReturn = MsgBox("There is a script scheduled to be executed, closing this window will cancel the script, continue?", vbOKCancel, "Script Control")
        If lngReturn = vbOK Then
            'Cancel the script
            Call sceScriptEngine.CancelAlarm
        Else
            Cancel = 1
        End If
    End If
Form_QueryUnload_ErrorHandler:
    Exit Sub
End Sub

Private Sub txtScript_Change()
'Created by Robin G Brown, 22/7/97
'Sub Declares
    'Error Trap
    On Error Resume Next
    booScriptChanged = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Created by Robin G Brown, 7/7/97
'Unload the form and clear the script engine object
'Function Declares
Const conFunction = "Form_Unload"
Dim lngTempValue As Long
    'Errors 'disabled' to ensure that form unloads
    On Error Resume Next
    Set sceScriptEngine = Nothing
    'SHOW_CALLING_FORM_IF_REQUIRED
End Sub


