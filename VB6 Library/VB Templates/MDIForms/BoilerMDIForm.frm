VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Begin VB.MDIForm frmMDIBoilerplate 
   BackColor       =   &H8000000C&
   Caption         =   "MDIBoilerplate"
   ClientHeight    =   6330
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11310
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.StatusBar stbStatus 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   6015
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   14314
            Key             =   "panStatus"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "panOther"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "9/2/97"
            Key             =   "panDateTime"
            Object.Tag             =   ""
         EndProperty
      EndProperty
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
         Caption         =   "Close"
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   4
      End
   End
   Begin VB.Menu mnuWindowRoot 
      Caption         =   "Window"
      Begin VB.Menu mnuWindow 
         Caption         =   "Tile &Horizontal"
         Index           =   1
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "Tile &Vertical"
         Index           =   2
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "&Cascade"
         Index           =   3
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "&Arrange Icons"
         Index           =   4
      End
      Begin VB.Menu mnuWindowSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowList 
         Caption         =   "List"
         WindowList      =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMDIBoilerplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Robin G Brown, INSERT_LONG_DATE
'-----------------------------------------------------------------------------
'   This form contains code for _
    DESCRIPTION_OF_MDI_FORM_FUNCTIONALITY_HERE
'-----------------------------------------------------------------------------
'   SPECIAL_INSTRUCTIONS
'   1. Fill in DESCRIPTION_OF_MDI_FORM_FUNCTIONALITY_HERE above
'   2. Replace all INSERT_DATE with date of creation
'   4. Rename form in Properties window (F4)
'   5. Replace all frmMDIBoilerplate with name of form
'   6. Replace all frmMDIBoilerplate with name of MDI parent form in frmMDIChildBoilerPlate forms
'   7. Fill in all DEFAULT_BEHAVIOUR
'   8. Delete this SPECIAL_INSTRUCTIONS set of comments
'   9. Save File As... with suitable filename
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "frmMDIBoilerplate"
Private intCounter As Integer
Private intReturn As Integer

Private Sub MDIForm_Activate()
'Created by Robin G Brown, INSERT_DATE
'DEFAULT_BEHAVIOUR
'Sub Declares
Const conSub = "Form_Activate"
Dim lngTempValue As Long
    'Error Trap
    On Error GoTo Form_Activate_ErrorHandler
    'CODE_HERE
Exit Sub
'Error Handler
Form_Activate_ErrorHandler:
    Select Case Err
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Error$(Err), vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Private Sub MDIForm_Load()
'Created by Robin G Brown, INSERT_DATE
'DEFAULT_BEHAVIOUR
'Sub Declares
Const conSub = "Form_Load"
    'Error Trap
    On Error GoTo Form_Load_ErrorHandler
    'CODE_HERE
Exit Sub
'Error Handler
Form_Load_ErrorHandler:
    Select Case Err
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Error$(Err), vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Private Sub MDIForm_Resize()
'Created by Robin G Brown, INSERT_DATE
'DEFAULT_BEHAVIOUR
'Sub Declares
Const conSub = "Form_Resize"
    'Error Trap
    On Error GoTo Form_Resize_ErrorHandler
    'Code to centre screen on window (delete as applicable)
    If Me.WindowState = 0 Then
        Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    End If
    'CODE_HERE
Exit Sub
'Error Handler
Form_Resize_ErrorHandler:
    Select Case Err
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Error$(Err), vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
'Created by Robin G Brown, INSERT_DATE
'DEFAULT_BEHAVIOUR
'Function Declares
Const conFunction = "Form_Unload"
Dim lngTempValue As Long
    'Errors 'disabled' to ensure that form unloads
    On Error Resume Next
    'CODE_HERE
End Sub

Private Sub mnuFile_Click(Index As Integer)
'Created by Robin G Brown, INSERT_DATE
'DEFAULT_BEHAVIOUR
    'Error Trap
    On Error Resume Next
    Select Case Index
    Case 1
        Call frmMDIChildBoilerPlate.OpenfrmMDIChildBoilerPlate
    Case 2
        Call frmMDIChildBoilerPlate.SavefrmMDIChildBoilerPlate
    Case 3
        Unload Me.ActiveForm
    Case 4
        Unload Me
    Case Else
        'Ignore
    End Select
End Sub

Private Sub mnuWindow_Click(Index As Integer)
'Created by Robin G Brown, INSERT_DATE
'Arrange child windows
    'Error Trap
    On Error Resume Next
    Select Case Index
    Case 1
        Me.Arrange vbTileHorizontal
    Case 2
        Me.Arrange vbTileVertical
    Case 3
        Me.Arrange vbCascade
    Case 4
        Me.Arrange vbArrangeIcons
    Case Else
        'Ignore
    End Select
End Sub

Private Sub stbStatus_PanelClick(ByVal Panel As ComctlLib.Panel)
'Created by Robin G Brown, INSERT_DATE
'DEFAULT_BEHAVIOUR
'Sub Declares
    'Error Trap
    On Error Resume Next
    With Panel
        Select Case .Key
        Case "panDateTime"
            If .Style = sbrDate Then
                .Style = sbrTime
            Else
                .Style = sbrDate
            End If
        Case "panStatus"
            .Tag = ""
            .Text = ""
            .ToolTipText = ""
        Case "panOther"
            'CODE_HERE
        Case Else
            'do nothing
        End Select
    End With
End Sub

Public Sub SetStatus(ByVal strStatus As String)
'Created by Robin G Brown, INSERT_DATE
'DEFAULT_BEHAVIOUR
'Sub Declares
    'Error Trap
    On Error Resume Next
    'Set the value of the status bar
    With stbStatus.Panels("status")
        .Tag = strStatus
        .Text = strStatus
        .ToolTipText = strStatus
    End With
End Sub

