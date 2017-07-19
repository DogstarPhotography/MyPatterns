VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "comdlg32.ocx"
Begin VB.Form frmCSVViewer 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "CSV Viewer"
   ClientHeight    =   3630
   ClientLeft      =   1155
   ClientTop       =   1845
   ClientWidth     =   5625
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3630
   ScaleWidth      =   5625
   Begin MSComDlg.CommonDialog cdlFile 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin MSFlexGridLib.MSFlexGrid grdDisplay 
      Height          =   3495
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   6165
      _Version        =   65541
   End
   Begin VB.Menu mnuFileRoot 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&Open"
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Save"
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmCSVViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Robin G Brown, 5th September 1997
'-----------------------------------------------------------------------------
'   This form contains code for _
    viewing CSV files
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "frmCSVViewer"
Private lngCounter As Long
Private lngReturn As Long
Private cdbCSVDatabase As CSVBase

Private Sub Form_Load()
'Created by Robin G Brown, 5/9/97
'Default behaviour
'Sub Declares
Const conSub = "Form_Load"
    'Error Trap
    On Error GoTo Form_Load_ErrorHandler
    Set cdbCSVDatabase = New CSVBase
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

Private Sub Form_Resize()
'Created by Robin G Brown, 10/9/97
'Sub Declares
Const conSub = "Form_Resize"
    'Error Trap
    On Error Resume Next
    If Me.WindowState <> 1 Then
        grdDisplay.Height = Me.Height - 820
        grdDisplay.Width = Me.Width - 240
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Created by Robin G Brown, 5/9/97
'Default behaviour
'Function Declares
Const conSub = "Form_Unload"
    'Errors 'disabled' to ensure that form unloads
    On Error Resume Next
    Set cdbCSVDatabase = Nothing
End Sub

Private Sub mnuFile_Click(Index As Integer)
'Created by Robin G Brown, 10/9/97
'Handle user requests
'Sub Declares
Const conSub = "mnuFile_Click"
Dim strFileName As String
Dim strDirectory As String
Dim strTable As String
Dim ctbTable As CSVTable
    'Error Trap
    On Error GoTo mnuFile_Click_ErrorHandler
    'Call the selected routine passing a reference to the current form
    Select Case Index
    Case 1
        'File Open
        With cdlFile
            .CancelError = True
            .Filter = conAnyCSVFilter
            .FilterIndex = 1
            .ShowOpen
            strFileName = .filename
        End With
        'Set the root directory
        strDirectory = GetPathFromPathAndFilename(strFileName)
        cdbCSVDatabase.Directory = strDirectory
        'Add the selected table
        strTable = StripFileExtension(GetFileNameFromPath(strFileName))
        cdbCSVDatabase.AddTable strTable
        Set ctbTable = cdbCSVDatabase.ctbTables(strTable)
        ctbTable.DisplayInGrid grdDisplay
        Me.Caption = strFileName
    Case 2
        'File Save
        strTable = StripFileExtension(GetFileNameFromPath(Me.Caption))
        Set ctbTable = cdbCSVDatabase.ctbTables(strTable)
        ctbTable.FileSave
        Me.Caption = "CSV Viewer"
        grdDisplay.Clear
        grdDisplay.FixedRows = 1
        grdDisplay.FixedCols = 1
        grdDisplay.Rows = 2
        grdDisplay.Cols = 2
    Case 3
        'Exit
        Unload Me
    Case Else
        'Default - do nothing
    End Select
Exit Sub
mnuFile_Click_ErrorHandler:
    MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    Exit Sub
End Sub
