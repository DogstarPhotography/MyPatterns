VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form mdiTreeChild 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Tree"
   ClientHeight    =   8145
   ClientLeft      =   1065
   ClientTop       =   1425
   ClientWidth     =   2370
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
   Icon            =   "TreeChild.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8145
   ScaleWidth      =   2370
   Begin ComctlLib.TreeView tvwScratch 
      DragIcon        =   "TreeChild.frx":0442
      Height          =   3255
      Left            =   0
      TabIndex        =   1
      Top             =   4860
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   5741
      _Version        =   327680
      Indentation     =   176
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "TreeChild.frx":0884
   End
   Begin ComctlLib.TreeView tvwSource 
      DragIcon        =   "TreeChild.frx":08A0
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   8493
      _Version        =   327680
      HideSelection   =   0   'False
      Indentation     =   176
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "TreeChild.frx":0CE2
   End
End
Attribute VB_Name = "mdiTreeChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Robin G Brown, 22nd April 1997
'-----------------------------------------------------------------------------
'   This form contains code for _
    displaying the source tree
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "mdiTreeChild"
Private intCounter As Integer
Private intReturn As Integer
Private Sub Form_Load()
'Created by Robin G Brown, 22/4/97
'Default behaviour
'Sub Declares
Const conSub = "Form_Load"
Dim nodtemp As Node
    'Error Trap
    On Error GoTo Form_Load_ErrorHandler
    tvwSource.ImageList = frmMaster.imlNode
    tvwScratch.ImageList = frmMaster.imlNode
    Call PrepareData
    Call RetrieveTreeSettings
    Me.Width = 3000
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
'Created by Robin G Brown, 22/4/97
'Default behaviour
'Sub Declares
Const conSub = "Form_Resize"
    'Error Trap
    On Error GoTo Form_Resize_ErrorHandler
    If Me.WindowState <> 1 Then
        If Me.Height < 1000 Then Me.Height = 1000
        If Me.Width < 1000 Then Me.Width = 1000
        Call ArrangeSourceAndScratch
    End If
Exit Sub
'Error Handler
Form_Resize_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub
Public Sub ArrangeSourceAndScratch()
'Created by Robin G Brown, 25/4/97
'Arrange the source and scratchpad treeview controls
'Sub Declares
    'Error Trap
    On Error Resume Next
    If frmMaster.mnuWindowScratchpad.Checked = True Then
        tvwSource.Height = (Me.Height / 3) * 2
        tvwSource.Width = Me.Width - 120
        tvwScratch.Visible = True
        tvwScratch.Top = tvwSource.Top + tvwSource.Height + 30
        tvwScratch.Height = Me.Height / 3 - 150
        tvwScratch.Width = Me.Width - 120
    Else
        tvwSource.Height = Me.Height - 120
        tvwSource.Width = Me.Width - 120
        tvwScratch.Visible = False
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Created by Robin G Brown, 25/4/97
'do not allow the child to be closed unless the program is ending
'Function Declares
Const conFunction = "Form_Unload"
Dim lngTempValue As Long
    'Errors 'disabled' to ensure that form unloads
    On Error Resume Next
    If UnloadMode = vbFormControlMenu Then
        MsgBox "You may not close this window."
        Cancel = True
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
'Created by Robin G Brown, 22/4/97
'Default behaviour
'Function Declares
Const conFunction = "Form_Unload"
Dim lngTempValue As Long
    'Errors 'disabled' to ensure that form unloads
    On Error Resume Next
    'CODE_HERE
End Sub
Private Sub tvwSource_DragDrop(Source As Control, X As Single, y As Single)
'Created by Robin G Brown, 15/4/97
'Handles moving nodes around using dragdrop functionality
'Sub Declares
Const conSub = "tvwSource_DragDrop"
    On Error Resume Next
    'Delete any nodes dropped onto this control, unless they are from this control
    If Not tvwSource Is Source Then
        'Do not delete any node without a parent as it is probably a root node
        If Not nodDrag.Parent Is Nothing Then
            Source.Nodes.Remove nodDrag.Key
        End If
    End If
    Set tvwSource.DropHighlight = Nothing
    Set Source.DropHighlight = Nothing
    tvwSource.Refresh
    Source.Refresh
    booDrag = False
    StatusBar ""
End Sub
Private Sub tvwSource_GotFocus()
    Set tvwCurrent = tvwSource
End Sub
Private Sub tvwSource_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    'Select the node
    Set nodDrag = tvwSource.HitTest(X, y)
    Set tvwCurrent = tvwSource
    If Not nodDrag Is Nothing Then
        nodDrag.Selected = True
        'Now popup a menu if the right button was pressed
        If Button = vbRightButton Then
            If Not tvwCurrent Is Nothing Then
                frmMaster.PopupMenu frmMaster.mnuTree
            End If
        End If
    End If
End Sub
Private Sub tvwSource_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    'Start the dragdrop operation
    On Error Resume Next
    If Button = vbLeftButton Then
        'Start drag unless the current node has no parent, _
            in which case it is probably the root node
        If Not nodDrag.Parent Is Nothing Then
            booDrag = True
            tvwSource.Drag vbBeginDrag
        End If
    End If
End Sub
Private Sub tvwScratch_DragDrop(Source As Control, X As Single, y As Single)
'Created by Robin G Brown, 15/4/97
'Handles moving nodes around using dragdrop functionality
'Sub Declares
Const conSub = "tvwScratch_DragDrop"
    On Error GoTo tvwScratch_DragDrop_ErrorHandler
    'Drop the source onto the target
    If Not tvwScratch.DropHighlight Is Nothing Then
        'If the target is not the source
        If Not nodDrag = tvwScratch.DropHighlight Then
            If tvwScratch Is Source Then
                'If we are moving within the same control, just change the parent
                Set nodDrag.Parent = tvwScratch.DropHighlight
            Else
                'Validate the action, as some movements are not allowed, _
                    assume that we are dropping onto one of the 'target' controls
                'No validation required for loading the scratchpad
                'Copy whole branch to target and delete source node/branch, unless it came from tvwScratch(0)
                Call CopyBranch(tvwScratch, nodDrag, tvwScratch.DropHighlight.Key)
                'Delete source branch if it is not from tvwSource
                If Not Source Is tvwSource Then
                    Source.Nodes.Remove nodDrag.Key
                End If
            End If
        End If
    End If
    'Finish off but don't throw an error
    On Error Resume Next
    'Show the dropped nodes in some way
    If nodDrag.Expanded = True Then
        'Call ExpandBranch(tvwScratch, tvwScratch.DropHighlight.Key)
        tvwScratch.DropHighlight.Child.EnsureVisible
    Else
        tvwScratch.DropHighlight.Child.EnsureVisible
    End If
    Set tvwScratch.DropHighlight = Nothing
    Set Source.DropHighlight = Nothing
    tvwScratch.Refresh
    Source.Refresh
    booDrag = False
    StatusBar ""
Exit Sub
tvwScratch_DragDrop_ErrorHandler:
    If Err = 55555 Then
        MsgBox "That action is not allowed."
    Else
        MsgBox "Invalid operation, please reset"
    End If
    On Error Resume Next
    Set tvwScratch.DropHighlight = Nothing
    Set Source.DropHighlight = Nothing
    tvwScratch.Refresh
    Source.Refresh
    booDrag = False
    Exit Sub
End Sub
Private Sub tvwScratch_DragOver(Source As Control, X As Single, y As Single, State As Integer)
    If booDrag = True Then
        'Set DropHighlight to the mouse's coordinates
        Set tvwScratch.DropHighlight = tvwScratch.HitTest(X, y)
        If Not tvwScratch.DropHighlight Is Nothing Then
            StatusBar "Release the mouse button to drop onto : " & tvwScratch.DropHighlight.Text
        End If
    End If
End Sub
Private Sub tvwScratch_GotFocus()
    Set tvwCurrent = tvwScratch
End Sub
Private Sub tvwScratch_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    'Select the node
    Set nodDrag = tvwScratch.HitTest(X, y)
    Set tvwCurrent = tvwScratch
    If Not nodDrag Is Nothing Then
        nodDrag.Selected = True
        'Now popup a menu if the right button was pressed
        If Button = vbRightButton Then
            If Not tvwCurrent Is Nothing Then
                'Me.PopupMenu mnuPopup
            End If
        End If
    End If
End Sub
Private Sub tvwScratch_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    'Start the dragdrop operation
    On Error Resume Next
    If Button = vbLeftButton Then
        If nodDrag.Parent Is Nothing Then
            'Do not allow drag operation, i.e. do nothing
        Else
        If nodDrag.Parent Is Nothing Then
            'Do not allow drag operation, i.e. do nothing
        Else
            booDrag = True
            tvwScratch.Drag vbBeginDrag
        End If
        End If
    End If
End Sub
Public Sub ChangeImages()
'Created by Robin G Brown, 25/4/97
'Change the image lists associated with the controls and refresh
'Sub Declares
Dim nodCurrent As Node
    'Error Trap
    'On Error Resume Next
    tvwSource.ImageList = frmMaster.imlNode
    For Each nodCurrent In tvwSource.Nodes
        nodCurrent.Image = "default"
    Next
    tvwSource.Indentation = 100
    tvwScratch.ImageList = frmMaster.imlNode
    For Each nodCurrent In tvwScratch.Nodes
        nodCurrent.Image = "default"
    Next
    tvwScratch.Indentation = 100
End Sub
Public Sub DefaultImages()
'Created by Robin G Brown, 29/4/97
'Change the image lists associated with the controls but don't refresh
'Sub Declares
Dim nodCurrent As Node
Dim strImage() As String
Dim intNodecounter As Integer
    'Error Trap
    On Error Resume Next
    tvwSource.ImageList = frmMaster.imlDefault
    tvwScratch.ImageList = frmMaster.imlDefault
End Sub
Public Sub PrepareData()
'Created by Robin G Brown, 21/4/97
'Fill the treeview controls with some data for testing
'Sub Declares
Const conSub = "PrepareData"
Dim intMemberCounter As Integer
Dim strMember As String
Dim intSubCounter As Integer
Dim strSubMember As String
Dim nodtemp As Node
    On Error GoTo PrepareData_ErrorHandler
    'First of all clear the controls
    tvwSource.Nodes.Clear
    'Prepare the 'Source' control
    'Add a root
    Set nodtemp = tvwSource.Nodes.Add(, , "root", "All Dimensions", conDefaultImage)
    nodtemp.EnsureVisible
    'Add the time dimension
    Set nodtemp = tvwSource.Nodes.Add( _
        "root", _
        tvwChild, _
        "time", _
        "Time", conDefaultImage)
    nodtemp.EnsureVisible
    'Add some years
    For intMemberCounter = 1990 To 1996
        'Add a branch
        strMember = CStr(intMemberCounter) & "_"
        Set nodtemp = tvwSource.Nodes.Add("time", tvwChild, "time_" & LCase(strMember), strMember, conDefaultImage)
        For intSubCounter = 1 To 12
            'Add a number of months
            strSubMember = strMember & Left$( _
                Choose(intSubCounter, "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December") _
                , 3)
            Set nodtemp = tvwSource.Nodes.Add("time_" & LCase(strMember), tvwChild, "time_" & LCase(strSubMember), strSubMember, conDefaultImage)
            'nodtemp.EnsureVisible
        Next
    Next
    'Add the accounts dimension
    Set nodtemp = tvwSource.Nodes.Add( _
        "root", _
        tvwChild, _
        "account", _
        "Accounts", conDefaultImage)
    nodtemp.EnsureVisible
    'Add some accounts
    For intMemberCounter = 1 To 6
        'Add a branch
        strMember = Choose(intMemberCounter, "Sales_Volume", "Sales_Revenue", "Sales_Expenses", "OverHead", "Margin", "Profit")
        Set nodtemp = tvwSource.Nodes.Add("account", tvwChild, "account_" & LCase(strMember), strMember, conDefaultImage)
        'nodtemp.EnsureVisible
    Next
    'Add the territory dimension
    Set nodtemp = tvwSource.Nodes.Add( _
        "root", _
        tvwChild, _
        "territory", _
        "Territories", conDefaultImage)
    nodtemp.EnsureVisible
    'Add some territories
    For intMemberCounter = 1 To 4
        'Add a branch
        strMember = Choose(intMemberCounter, "USA", "Europe", "Asia", "Other_Territories")
        Set nodtemp = tvwSource.Nodes.Add("territory", tvwChild, LCase(strMember), strMember, conDefaultImage)
        For intSubCounter = 1 To 4
            'Add a number of leaves
            strSubMember = Choose(intSubCounter, "North", "South", "East", "West") & "_" & strMember
            Set nodtemp = tvwSource.Nodes.Add(LCase(strMember), tvwChild, LCase(strSubMember), strSubMember, conDefaultImage)
            'nodtemp.EnsureVisible
        Next
    Next
    'Add the business class dimension
    Set nodtemp = tvwSource.Nodes.Add( _
        "root", _
        tvwChild, _
        "business", _
        "Business_Classes", conDefaultImage)
    nodtemp.EnsureVisible
    'Add some classes
    For intMemberCounter = 1 To 5
        strMember = Choose(intMemberCounter, "Maritime", "Vehicle", "Household", "Health", "Other_Business")
        Set nodtemp = tvwSource.Nodes.Add("business", tvwChild, LCase(strMember), strMember, conDefaultImage)
        For intSubCounter = 1 To 3
            strSubMember = strMember & "_" & Choose(intSubCounter, "A", "B", "C")
            Set nodtemp = tvwSource.Nodes.Add(LCase(strMember), tvwChild, LCase(strSubMember), strSubMember, conDefaultImage)
            'nodtemp.EnsureVisible
        Next
    Next
    'Add the sales team dimension
    Set nodtemp = tvwSource.Nodes.Add( _
        "root", _
        tvwChild, _
        "team", _
        "Sales_Teams", conDefaultImage)
    nodtemp.EnsureVisible
    For intMemberCounter = 1 To 4
        strMember = Choose(intMemberCounter, "John", "Paul", "George", "Ringo")
        Set nodtemp = tvwSource.Nodes.Add("team", tvwChild, LCase(strMember), strMember, conDefaultImage)
        For intSubCounter = 1 To 3
            strSubMember = strMember & "_" & Choose(intSubCounter, "Primary", "Secondary", "Tertiary")
            Set nodtemp = tvwSource.Nodes.Add(LCase(strMember), tvwChild, LCase(strSubMember), strSubMember, conDefaultImage)
            'nodtemp.EnsureVisible
        Next
    Next
    'Add the client dimension
    Set nodtemp = tvwSource.Nodes.Add( _
        "root", _
        tvwChild, _
        "client", _
        "Clients", conDefaultImage)
    nodtemp.EnsureVisible
    For intMemberCounter = 1 To 5
        strMember = Choose(intMemberCounter, "Client_A", "Client_B", "Client_C", "Client_D", "Client_E")
        Set nodtemp = tvwSource.Nodes.Add("client", tvwChild, LCase(strMember), strMember, conDefaultImage)
        For intSubCounter = 1 To 4
            strSubMember = strMember & "_" & Choose(intSubCounter, "Division_1", "Division_2", "Division_3", "Division_4")
            Set nodtemp = tvwSource.Nodes.Add(LCase(strMember), tvwChild, LCase(strSubMember), strSubMember, conDefaultImage)
            'nodtemp.EnsureVisible
        Next
    Next
    'Add a root to the scratchpad tree
    Set nodtemp = tvwScratch.Nodes.Add(, , "root", "Scratchpad", conDefaultImage)
Exit Sub
PrepareData_ErrorHandler:
    MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    Exit Sub
End Sub

Public Sub RetrieveTreeSettings()
'Created by Robin G Brown, 9/5/97
'Return the expanded state of the tree as a CSV string
'Function Declares
Const conFunction = "RetrieveTreeSettings"
Dim strCSV As String
Dim intCommaPos As Integer
Dim intLastPos As Integer
Dim strValue As String
    'Error Trap
    On Error GoTo RetrieveTreeSettings_ErrorHandler
    strCSV = strRegistrySettings(3)
    intLastPos = 1
    intCommaPos = InStr(intLastPos, strCSV, ",")
    intCounter = 1
    Do While intCommaPos > 0 And intCounter <= tvwSource.Nodes.Count
        strValue = Mid$(strCSV, intLastPos + 1, 1)
        If strValue = "E" Then
            tvwSource.Nodes(intCounter).Expanded = True
        Else
            tvwSource.Nodes(intCounter).Expanded = False
        End If
        intCounter = intCounter + 1
        intLastPos = intCommaPos
        intCommaPos = InStr(intLastPos + 1, strCSV, ",")
    Loop
    tvwSource.Nodes(1).Expanded = True
    tvwSource.Nodes(1).EnsureVisible
Exit Sub
'Error Handler
RetrieveTreeSettings_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conFunction
    End Select
    Exit Sub
End Sub

Public Sub GetTreeSettings()
'Created by Robin G Brown, 9/5/97
'Return the expanded state of the tree as a CSV string
'Function Declares
Const conFunction = "GetTreeSettings"
Dim strCSV As String
    'Error Trap
    On Error GoTo GetTreeSettings_ErrorHandler
    strCSV = ""
    For intCounter = 1 To tvwSource.Nodes.Count
        If tvwSource.Nodes(intCounter).Expanded = True Then
            strCSV = strCSV & "E,"
        Else
            strCSV = strCSV & "C,"
        End If
    Next
    strCSV = Left$(strCSV, Len(strCSV) - 1)
    strRegistrySettings(3) = strCSV
Exit Sub
'Error Handler
GetTreeSettings_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conFunction
    End Select
    strRegistrySettings(3) = ""
    Exit Sub
End Sub


