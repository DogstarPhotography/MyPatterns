VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "comdlg32.ocx"
Begin VB.Form frmMDIChildBoilerPlate 
   Caption         =   "MDIChildBoilerplate"
   ClientHeight    =   4455
   ClientLeft      =   1155
   ClientTop       =   1560
   ClientWidth     =   7305
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
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4455
   ScaleWidth      =   7305
   Begin MSComDlg.CommonDialog cdlfrmMDIChildBoilerPlate 
      Left            =   60
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
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
Attribute VB_Name = "frmMDIChildBoilerPlate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Robin G Brown, INSERT_LONG_DATE
'-----------------------------------------------------------------------------
'   This form contains code for _
    DESCRIPTION_OF_MDICHILD_FORM_FUNCTIONALITY_HERE
'-----------------------------------------------------------------------------
'   SPECIAL_INSTRUCTIONS
'   1. Fill in DESCRIPTION_OF_MDICHILD_FORM_FUNCTIONALITY_HERE above
'   2. Replace all INSERT_DATE with date of creation
'   4. Rename form in Properties window (F4)
'   5. Replace all MDIChildBoilerPlate with name of form (without prefix)
'   6. Replace all frmMDIBoilerplate with name of MDI parent form
'   7. Rename cdlMDIChildBoilerPlate
'   8. Fill in all DEFAULT_BEHAVIOUR
'   9. Delete this SPECIAL_INSTRUCTIONS set of comments
'   10. Save File As... with suitable filename
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "frmMDIChildBoilerPlate"
'Variables to store saved status of this form
Private booIsDirty As Boolean
Private strFileName As String
Private lngCounter As Long
Private lngReturn As Long

Private Sub Form_Activate()
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

Private Sub Form_Initialize()
'Created by Robin G Brown, INSERT_DATE
'DEFAULT_BEHAVIOUR
'Sub Declares
    'Error Trap
    On Error Resume Next
    Err.Clear
    'CODE_HERE
    booIsDirty = True
    Set PROPERTYOBJECT = Nothing
    If Err.Number <> 0 Then
        'ERROR_HANDLING_CODE_HERE
        Err.Raise Err.Number, conClass & ":Initialize", "Unexpected error : " & Err.Description
    End If
End Sub

Private Sub Form_Load()
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

Private Sub Form_Resize()
'Created by Robin G Brown, INSERT_DATE
'DEFAULT_BEHAVIOUR
'Sub Declares
Const conSub = "Form_Resize"
    'Error Trap
    On Error GoTo Form_Resize_ErrorHandler
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

Private Sub Form_Unload(Cancel As Integer)
'Created by Robin G Brown, INSERT_DATE
'DEFAULT_BEHAVIOUR
'Function Declares
Const conFunction = "Form_Unload"
Dim lngTempValue As Long
    'Errors 'disabled' to ensure that form unloads
    On Error Resume Next
    If booIsDirty = True Then
        lngReturn = MsgBox("frmMDIChildBoilerPlate not saved. Save before closing?", vbYesNoCancel, App.ProductName)
        Select Case lngReturn
        Case vbYes
            'Save
            Call SavefrmMDIChildBoilerPlate
        Case vbCancel
            'Cancel close
            Cancel = 1
        Case Else
            'Carry on without saving
        End Select
    End If
    'CODE_HERE
End Sub

Public Sub SavefrmMDIChildBoilerPlate()
'Created by Robin G Brown, INSERT_DATE
'DESCRIPTION_OF_SUBROUTINE_FUNCTIONALITY_HERE
'Sub Declares
Const conSub = "SavefrmMDIChildBoilerPlate"
    'Error Trap
    On Error GoTo SavefrmMDIChildBoilerPlate_ErrorHandler
    With cdlfrmMDIChildBoilerPlate
        .DialogTitle = "Save MDIChildBoilerPlate"
        .CancelError = True
        .Flags = cdlOFNHideReadOnly
        .Filter = conAnyFilter
        .FilterIndex = 1
        .ShowSave
        strFileName = .FileName
    End With
    'CODE_HERE
    booIsDirty = False
Exit Sub
'Error Handler
SavefrmMDIChildBoilerPlate_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Public Sub OpenfrmMDIChildBoilerPlate()
'Created by Robin G Brown, INSERT_DATE
'DESCRIPTION_OF_SUBROUTINE_FUNCTIONALITY_HERE
'Sub Declares
Const conSub = "OpenfrmMDIChildBoilerPlate"
    'Error Trap
    On Error GoTo OpenfrmMDIChildBoilerPlate_ErrorHandler
    'CODE_HERE
    With cdlfrmMDIChildBoilerPlate
        .DialogTitle = "open MDIChildBoilerPlate"
        .CancelError = True
        .Flags = cdlOFNHideReadOnly
        .Filter = conAnyFilter
        .FilterIndex = 1
        .ShowOpen
        strFileName = .FileName
    End With
    booIsDirty = False
Exit Sub
'Error Handler
OpenfrmMDIChildBoilerPlate_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Private Sub mnuFile_Click(Index As Integer)
'Created by Robin G Brown, INSERT_DATE
'DEFAULT_BEHAVIOUR
    'Error Trap
    On Error Resume Next
    Select Case Index
    Case 1
        Call Me.OpenfrmMDIChildBoilerPlate
    Case 2
        Call Me.SavefrmMDIChildBoilerPlate
    Case 3
        Unload Me
    Case 4
        Unload frmMDIBoilerplate
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
        frmMDIBoilerplate.Arrange vbTileHorizontal
    Case 2
        frmMDIBoilerplate.Arrange vbTileVertical
    Case 3
        frmMDIBoilerplate.Arrange vbCascade
    Case 4
        frmMDIBoilerplate.Arrange vbArrangeIcons
    Case Else
        'Ignore
    End Select
End Sub

'-----------------------------------------------------------------------------
'   Properties
'-----------------------------------------------------------------------------

Property Get IsDirty() As Boolean
'Created by Robin G Brown, INSERT_DATE
'Property Declares
Const conProperty = "IsDirty"
'Error Trap
    On Error Resume Next
    IsDirty = booIsDirty
End Property

Property Get FileName() As Boolean
'Created by Robin G Brown, INSERT_DATE
'Property Declares
Const conProperty = "FileName"
'Error Trap
    On Error Resume Next
    FileName = strFileName
End Property
