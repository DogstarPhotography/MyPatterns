VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "comdlg32.ocx"
Begin VB.Form frmDisplaySettings 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Account Display Settings"
   ClientHeight    =   5145
   ClientLeft      =   1155
   ClientTop       =   1560
   ClientWidth     =   8010
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
   ScaleHeight     =   5145
   ScaleWidth      =   8010
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgAccounts 
      Left            =   6720
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin VB.Frame fraAccount 
      Caption         =   "IBNR"
      Height          =   1035
      Index           =   4
      Left            =   60
      TabIndex        =   14
      Top             =   4020
      Width           =   6435
      Begin VB.CheckBox chkAccount 
         Caption         =   "Display as %"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   33
         Top             =   600
         Width           =   1875
      End
      Begin VB.ComboBox cmbMask 
         Height          =   315
         Index           =   9
         Left            =   4140
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtAccount 
         Height          =   285
         Index           =   9
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   28
         Top             =   600
         Width           =   1995
      End
      Begin VB.CheckBox chkAccount 
         Caption         =   "Display"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1275
      End
      Begin VB.TextBox txtAccount 
         Height          =   285
         Index           =   8
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   16
         Top             =   240
         Width           =   1995
      End
      Begin VB.ComboBox cmbMask 
         Height          =   315
         Index           =   8
         Left            =   4140
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame fraAccount 
      Caption         =   "Incurred"
      Height          =   1035
      Index           =   3
      Left            =   60
      TabIndex        =   11
      Top             =   2940
      Width           =   6435
      Begin VB.CheckBox chkAccount 
         Caption         =   "Display as %"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   1875
      End
      Begin VB.ComboBox cmbMask 
         Height          =   315
         Index           =   7
         Left            =   4140
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtAccount 
         Height          =   285
         Index           =   7
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   26
         Top             =   600
         Width           =   1995
      End
      Begin VB.CheckBox chkAccount 
         Caption         =   "Display"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1275
      End
      Begin VB.TextBox txtAccount 
         Height          =   285
         Index           =   6
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   13
         Top             =   240
         Width           =   1995
      End
      Begin VB.ComboBox cmbMask 
         Height          =   315
         Index           =   6
         Left            =   4140
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   240
         Width           =   2115
      End
   End
   Begin VB.Frame fraAccount 
      Caption         =   "O/S Claim"
      Height          =   1035
      Index           =   2
      Left            =   60
      TabIndex        =   8
      Top             =   1860
      Width           =   6435
      Begin VB.CheckBox chkAccount 
         Caption         =   "Display as %"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   1875
      End
      Begin VB.ComboBox cmbMask 
         Height          =   315
         Index           =   5
         Left            =   4140
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtAccount 
         Height          =   285
         Index           =   5
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   24
         Top             =   600
         Width           =   1995
      End
      Begin VB.CheckBox chkAccount 
         Caption         =   "Display"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1275
      End
      Begin VB.TextBox txtAccount 
         Height          =   285
         Index           =   4
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   10
         Top             =   240
         Width           =   1995
      End
      Begin VB.ComboBox cmbMask 
         Height          =   315
         Index           =   4
         Left            =   4140
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame fraAccount 
      Caption         =   "Paid Claim"
      Height          =   1035
      Index           =   1
      Left            =   60
      TabIndex        =   5
      Top             =   780
      Width           =   6435
      Begin VB.CheckBox chkAccount 
         Caption         =   "Display as %"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   30
         Top             =   600
         Width           =   1875
      End
      Begin VB.ComboBox cmbMask 
         Height          =   315
         Index           =   3
         Left            =   4140
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtAccount 
         Height          =   285
         Index           =   3
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   22
         Top             =   600
         Width           =   1995
      End
      Begin VB.CheckBox chkAccount 
         Caption         =   "Display"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   300
         Width           =   1815
      End
      Begin VB.TextBox txtAccount 
         Height          =   285
         Index           =   2
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   7
         Top             =   240
         Width           =   1995
      End
      Begin VB.ComboBox cmbMask 
         Height          =   315
         Index           =   2
         Left            =   4140
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   6600
      TabIndex        =   2
      Top             =   540
      Width           =   1335
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "OK"
      Height          =   375
      Index           =   0
      Left            =   6600
      TabIndex        =   1
      Top             =   60
      Width           =   1335
   End
   Begin VB.Frame fraAccount 
      Caption         =   "Premium"
      Height          =   675
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6435
      Begin VB.CheckBox chkAccount 
         Caption         =   "Display"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmbMask 
         Height          =   315
         Index           =   1
         Left            =   4140
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtAccount 
         Height          =   285
         Index           =   1
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   3
         Top             =   240
         Width           =   1995
      End
   End
End
Attribute VB_Name = "frmDisplaySettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Robin G Brown, 7th May 1997
'-----------------------------------------------------------------------------
'   This form contains code for _
    changing display settings for frmGridReport
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "frmDisplaySettings"
Private intCounter As Integer
Private intReturn As Integer
Private dtaDataArray As TriData

Property Set DataArray(ByRef objDataArray As Object)
'Created by Robin G Brown, 19/8/97
'Set the dtaDataArray property
'Sub Declares
Const conProperty = "DataArray"
    'Error Trap
    On Error Resume Next
    Set dtaDataArray = objDataArray
    If Err.Number <> 0 Then
        Err.Raise Err.Number, conModule & ":" & conProperty, "Unexpected error : " & Err.Description
    End If
End Property

Private Sub cmdAction_Click(Index As Integer)
'Created by Robin G Brown, 7/5/97
'Respond to user action
'Sub Declares
Const conSub = "cmdAction_Click"
Dim intAccount As Integer
    'Error Trap
    On Error GoTo cmdAction_Click_ErrorHandler
    Select Case Index
    Case 0
        'OK - Save settings
        'Display columns first
        For intCounter = 1 To 9
            If chkAccount(intCounter).Value = 1 Then
                Call dtaDataArray.ShowAccount(intCounter, True)
            Else
                Call dtaDataArray.ShowAccount(intCounter, False)
            End If
            'Now check new names
            intAccount = intCounter * 2
            dtaDataArray.AccountName(intCounter) = txtAccount(intCounter).Text
            intAccount = intCounter * 2
            'Lastly apply format masks
            intAccount = intCounter * 2
            If cmbMask(intCounter).ListIndex > 0 Then
                Select Case cmbMask(intCounter).Text
                Case "Standard"
                    dtaDataArray.AccountMask(intCounter) = "0"
                Case "Fixed '0.00'"
                    dtaDataArray.AccountMask(intCounter) = "0.00"
                Case "Percent"
                    dtaDataArray.AccountMask(intCounter) = "0\%"
                Case "Percent '0.0'"
                    dtaDataArray.AccountMask(intCounter) = "0.0\%"
                Case "Percent '0.00'"
                    dtaDataArray.AccountMask(intCounter) = "0.00\%"
                Case Else
                    dtaDataArray.AccountMask(intCounter) = "0"
                End Select
            End If
        Next
        'Then unload
        Unload Me
    Case 1
        'Cancel
        Unload Me
    Case Else
        'Do nothing
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

Private Sub Form_Load()
'Created by Robin G Brown, 7/5/97
'Default Behaviour
'Sub Declares
Const conSub = "Form_Load"
Dim intAccount As Integer
    'Error Trap
    On Error GoTo Form_Load_ErrorHandler
    'Set up the various controls
    'Set up various account checkboxes, names, colours, etc.
    For intCounter = 1 To 9
        'Set checkboxes
        If dtaDataArray.DisplayAccount(intCounter) = True Then
            chkAccount(intCounter).Value = 1
        End If
        'Set account name
        txtAccount(intCounter).Text = dtaDataArray.AccountName(intCounter)
        'Set colour
'        txtAccount(intCounter).ForeColor = dtaDataArray.AccountColour(intCounter)
        'txtAccount(intCounter).BackColor
        'Set formats
        Select Case dtaDataArray.AccountMask(intCounter)
        Case "0"
            cmbMask(intCounter).AddItem "[Standard]"
        Case "0.00"
            cmbMask(intCounter).AddItem "[Fixed '0.00']"
        Case "0\%"
            cmbMask(intCounter).AddItem "[Percent]"
        Case "0.0\%"
            cmbMask(intCounter).AddItem "[Percent '0.0']"
        Case "0.00\%"
            cmbMask(intCounter).AddItem "[Percent '0.00']"
        Case Else
            cmbMask(intCounter).AddItem "[" & dtaDataArray.AccountMask(intCounter) & "]"
        End Select
        If InStr(txtAccount(intCounter).Text, "%") Then
            cmbMask(intCounter).AddItem "Percent"
            cmbMask(intCounter).AddItem "Percent '0.0'"
            cmbMask(intCounter).AddItem "Percent '0.00'"
            cmbMask(intCounter).ListIndex = 0
        Else
            cmbMask(intCounter).AddItem "Standard"
            cmbMask(intCounter).AddItem "Fixed '0.00'"
            cmbMask(intCounter).ListIndex = 0
        End If
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
