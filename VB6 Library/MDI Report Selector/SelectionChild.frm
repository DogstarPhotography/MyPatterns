VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form mdiSelectionChild 
   Caption         =   "Selection"
   ClientHeight    =   4380
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   5355
   Icon            =   "SelectionChild.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4380
   ScaleWidth      =   5355
   Begin ComctlLib.TreeView tvwSelection 
      DragIcon        =   "SelectionChild.frx":0442
      Height          =   2715
      Index           =   1
      Left            =   60
      TabIndex        =   1
      Top             =   360
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   4789
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
      MouseIcon       =   "SelectionChild.frx":0884
   End
   Begin ComctlLib.TabStrip tabDimensions 
      Height          =   3555
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   6271
      ShowTips        =   0   'False
      _Version        =   327680
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "1"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "SelectionChild.frx":08A0
   End
End
Attribute VB_Name = "mdiSelectionChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Robin G Brown, 25th April 1997
'-----------------------------------------------------------------------------
'   This form contains code for _
    DESCRIPTION_OF_FORM_FUNCTIONALITY_HERE
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "mdiSelectionChild"
Private intCounter As Integer
Private intReturn As Integer
Private intDimensions As Integer
Private intDimensionImage As Integer
Private Sub Form_Load()
'Created by Robin G Brown, 25/4/97
'Populate the treeview controls
'Sub Declares
Const conSub = "Form_Load"
Dim nodtemp As Node
Dim imlTemp As ImageList
    'Error Trap
    On Error GoTo Form_Load_ErrorHandler
    'The only difference between forms is the icon on the tabs!
    Select Case Forms.Count - 2
    Case 1
        intDimensionImage = 1
        Me.Caption = "Height"
    Case 2
        intDimensionImage = 2
        Me.Caption = "Width"
    Case 3
        intDimensionImage = 3
        Me.Caption = "Depth"
    Case Else
        intDimensionImage = 1
        Me.Caption = "Unknown"
    End Select
    'Add some treeview controls
    tabDimensions.ImageList = frmMaster.imlDimension
    tabDimensions.Tabs(1).Image = intDimensionImage
    intDimensions = 1
    For intCounter = 1 To intDimensions
        If intCounter > 1 Then
            tabDimensions.Tabs.Add intCounter, "col" & intCounter, CStr(intCounter), intDimensionImage
            'tabDimensions.ZOrder conZBottom
            Load tvwSelection(intCounter)
        End If
        tvwSelection(intCounter).Visible = True
        'tvwSelection(intCounter).Visible = False
        'tvwSelection(intCounter).ZOrder conZTop
        tvwSelection(intCounter).ImageList = frmMaster.imlNode
        Set nodtemp = tvwSelection(intCounter).Nodes.Add(, , "root", "Column_" & intCounter, conDefaultImage)
    Next
    'tvwSelection(1).Visible = True
    'tvwSelection(1).ZOrder conZTop
    Me.Height = 3600
    Me.Width = 3600
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

Private Sub Form_Resize()
'Created by Robin G Brown, 25/4/97
'Resize controls on form
'Sub Declares
Const conSub = "Form_Resize"
    'Error Trap
    On Error GoTo Form_Resize_ErrorHandler
    If Me.WindowState <> 1 Then
        If Me.Height < 1000 Then Me.Height = 1000
        If Me.Width < 1000 Then Me.Width = 1000
        tabDimensions.Height = Me.Height - 420
        tabDimensions.Width = Me.Width - 120
        For intCounter = 1 To intDimensions
            tvwSelection(intCounter).Height = Me.Height - 820
            tvwSelection(intCounter).Width = Me.Width - 240
        Next
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
Private Sub Form_Unload(Cancel As Integer)
'Created by Robin G Brown, 25/4/97
'Default behaviour
'Function Declares
Const conFunction = "Form_Unload"
Dim lngTempValue As Long
    'Errors 'disabled' to ensure that form unloads
    On Error Resume Next
    'CODE_HERE
End Sub

Private Sub tabDimensions_Click()
'Created by Robin G Brown, 25/4/97
'Handle someone clicking on the tab, move treeview controls in ZOrder as appropriate
'Sub Declares
Const conSub = "ImmediateErrorSub"
    'Error Trap
    On Error GoTo ImmediateErrorSub_ErrorHandler
    'Indices on the tabstrip control match those for the treeview controls
    tvwSelection(tabDimensions.SelectedItem.Index).ZOrder conZTop
    'Make sure this is now the active form
    Me.SetFocus
Exit Sub
'Error Handler
ImmediateErrorSub_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Private Sub tabDimensions_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'Created by Robin G Brown, 25/4/97
'Pop up a menu if required
    'Error Trap
    On Error Resume Next
    'Now popup a menu if the right button was pressed
    If Button = vbRightButton Then
        'Make sure this is now the active form
        Me.SetFocus
        Me.PopupMenu frmMaster.mnuDimension
    End If
End Sub

Private Sub tvwSelection_DragDrop(Index As Integer, Source As Control, X As Single, y As Single)
'Created by Robin G Brown, 25/4/97
'Handles moving nodes around using dragdrop functionality
'Sub Declares
Const conSub = "tvwSelection_DragDrop"
    'Error Trap
    On Error GoTo tvwSelection_DragDrop_ErrorHandler
    'Set index to top control, to force correct selection of controls!
    Index = tabDimensions.SelectedItem.Index
    'Drop the source onto the target
    If Not tvwSelection(Index).DropHighlight Is Nothing Then
        'If the target is not the source
        If Not nodDrag = tvwSelection(Index).DropHighlight Then
            If tvwSelection(Index) Is Source Then
                'If we are moving within the same control, just change the parent
                Set nodDrag.Parent = tvwSelection(Index).DropHighlight
            Else
                'Validate the action, as some movements are not allowed, _
                    assume that we are dropping onto one of the 'target' controls
                'A series of validation questions, which must all be 'passed' in order to complete the operation
                'VALIDATION_CODE_HERE
                'Copy whole branch to target and delete source node/branch, unless it came from tvwSelection(0)
                Call CopyBranch(tvwSelection(Index), nodDrag, tvwSelection(Index).DropHighlight.Key)
                'Delete source branch if it is not from tvwSource
                If Not Source Is mdiTreeChild.tvwSource Then
                    Source.Nodes.Remove nodDrag.Key
                End If
            End If
        End If
    End If
    'Finish off but don't throw an error
    On Error Resume Next
    'Show the dropped nodes in some way
    If nodDrag.Expanded = True Then
        'Call ExpandBranch(tvwSelection(Index), tvwSelection(Index).DropHighlight.Key)
        tvwSelection(Index).DropHighlight.Child.EnsureVisible
    Else
        tvwSelection(Index).DropHighlight.Child.EnsureVisible
    End If
    Set tvwSelection(Index).DropHighlight = Nothing
    Set Source.DropHighlight = Nothing
    tvwSelection(Index).Refresh
    Source.Refresh
    booDrag = False
    StatusBar ""
Exit Sub
'Error Handler
tvwSelection_DragDrop_ErrorHandler:
    If Err = 55555 Then
        MsgBox "That action is not allowed."
    Else
        MsgBox "Invalid operation, please reset"
    End If
    On Error Resume Next
    Set tvwSelection(Index).DropHighlight = Nothing
    Set Source.DropHighlight = Nothing
    tvwSelection(Index).Refresh
    Source.Refresh
    booDrag = False
    Exit Sub
End Sub

Private Sub tvwSelection_DragOver(Index As Integer, Source As Control, X As Single, y As Single, State As Integer)
'Created by Robin G Brown, 25/4/97
'Default behaviour
'Sub Declares
Const conSub = "tvwSelection_DragOver"
    'Error Trap
    On Error GoTo tvwSelection_DragOver_ErrorHandler
    'Set index to top control, to force correct selection of controls!
    Index = tabDimensions.SelectedItem.Index
    If booDrag = True Then
        'Set DropHighlight to the mouse's coordinates
        Set tvwSelection(Index).DropHighlight = tvwSelection(Index).HitTest(X, y)
        If Not tvwSelection(Index).DropHighlight Is Nothing Then
            StatusBar "Release the mouse button to drop onto : " & tvwSelection(Index).DropHighlight.Text
        End If
    End If
Exit Sub
'Error Handler
tvwSelection_DragOver_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Private Sub tvwSelection_GotFocus(Index As Integer)
'Created by Robin G Brown, 25/4/97
'Default behaviour
'Sub Declares
Const conSub = "tvwSelection_GotFocus"
    'Error Trap
    On Error GoTo tvwSelection_GotFocus_ErrorHandler
    'Set index to top control, to force correct selection of controls!
    Index = tabDimensions.SelectedItem.Index
    Set tvwCurrent = tvwSelection(Index)
Exit Sub
'Error Handler
tvwSelection_GotFocus_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Private Sub tvwSelection_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
'Created by Robin G Brown, 25/4/97
'Default behaviour
'Sub Declares
Const conSub = "tvwSelection_MouseDown"
    'Error Trap
    On Error GoTo tvwSelection_MouseDown_ErrorHandler
    'Set index to top control, to force correct selection of controls!
    Index = tabDimensions.SelectedItem.Index
    'Select the node
    Set nodDrag = tvwSelection(Index).HitTest(X, y)
    Set tvwCurrent = tvwSelection(Index)
    If Not nodDrag Is Nothing Then
        nodDrag.Selected = True
        'Now popup a menu if the right button was pressed
        If Button = vbRightButton Then
            If Not tvwCurrent Is Nothing Then
                'Make sure this is now the active form
                Me.SetFocus
                Me.PopupMenu frmMaster.mnuTree
            End If
        End If
    End If
Exit Sub
'Error Handler
tvwSelection_MouseDown_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Private Sub tvwSelection_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
'Created by Robin G Brown, 25/4/97
'Default behaviour
'Sub Declares
Const conSub = "tvwSelection_MouseMove"
    'Error Trap
    On Error GoTo tvwSelection_MouseMove_ErrorHandler
    On Error Resume Next
    'Set index to top control, to force correct selection of controls!
    Index = tabDimensions.SelectedItem.Index
    'Start the dragdrop operation
    If Button = vbLeftButton Then
        If nodDrag.Parent Is Nothing Then
            'Do not allow drag operation on first node, i.e. do nothing
        Else
            booDrag = True
            tvwSelection(Index).Drag vbBeginDrag
        End If
    End If
Exit Sub
'Error Handler
tvwSelection_MouseMove_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub
Public Function GetDimensions() As Integer
'Created by Robin G Brown, 25/4/97
'return the number of dimensions
'Function Declares
Const conFunction = "GetDimensions"
Dim intDimCounter As Integer
    'Error Trap
    On Error GoTo GetDimensions_ErrorHandler
    intDimCounter = 0
    For intCounter = 1 To intDimensions
        If tvwSelection(intCounter).Nodes(1).Children > 0 Then
            intDimCounter = intDimCounter + 1
        End If
    Next
    GetDimensions = intDimCounter
Exit Function
'Error Handler
GetDimensions_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conFunction
    End Select
    GetDimensions = 0
    Exit Function
End Function
Public Sub ChangeImages()
'Created by Robin G Brown, 25/4/97
'Change the image lists associated with the controls and refresh
'Sub Declares
Dim nodCurrent As Node
Dim strImage() As String
Dim intNodecounter As Integer
    'Error Trap
    On Error Resume Next
    tabDimensions.ImageList = frmMaster.imlDimension
    For intCounter = 1 To intDimensions
        tabDimensions.Tabs(intCounter).Image = intDimensionImage
        'Read all the images first
        ReDim strImage(1 To 1)
        For intNodecounter = 1 To tvwSelection(intCounter).Nodes.Count
            ReDim Preserve strImage(1 To intNodecounter)
            strImage(intNodecounter) = tvwSelection(intCounter).Nodes(intNodecounter).Image
        Next
        tvwSelection(intCounter).ImageList = frmMaster.imlNode
        For intNodecounter = 1 To tvwSelection(intCounter).Nodes.Count
            tvwSelection(intCounter).Nodes(intNodecounter).Image = strImage(intNodecounter)
        Next
        tvwSelection(intCounter).Indentation = 100
    Next
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
    tabDimensions.ImageList = frmMaster.imlDefault
    'For intCounter = 1 To intDimensions
        'tabDimensions.Tabs(intCounter).Image = intDimensionImage
        'Read all the images first
        'ReDim strImage(1 To 1)
        'For intNodecounter = 1 To tvwSelection(intCounter).Nodes.Count
        '    ReDim Preserve strImage(1 To intNodecounter)
        '    strImage(intNodecounter) = tvwSelection(intCounter).Nodes(intNodecounter).Image
        'Next
        tvwSelection(intCounter).ImageList = frmMaster.imlDefault
        'For intNodecounter = 1 To tvwSelection(intCounter).Nodes.Count
        '    tvwSelection(intCounter).Nodes(intNodecounter).Image = strImage(intNodecounter)
        'Next
    'Next
End Sub

Public Sub AddDimension()
'Created by Robin G Brown, 25/4/97
'Adds a tab and a dimension to the form
'Sub Declares
Const conSub = "AddDimension"
Dim nodtemp As Node
    'Error Trap
    On Error GoTo AddDimension_ErrorHandler
    'Prevent more than 5 dimensions
    If intDimensions = 5 Then
        MsgBox "Cannot have more than 5 dimensions."
        Exit Sub
    End If
    'Warn when adding a new dimension if the prevous one is unused yet
    If tvwSelection(intDimensions).Nodes(1).Children = 0 Then
        If MsgBox("You are not using this dimension currently, add a new dimension anyway?", vbOKCancel) = vbCancel Then
            Exit Sub
        End If
    End If
    tabDimensions.Tabs.Add intDimensions + 1, "col" & intDimensions + 1, CStr(intDimensions + 1), intDimensionImage
    Load tvwSelection(intDimensions + 1)
    tvwSelection(intDimensions + 1).Visible = True
    tvwSelection(intDimensions + 1).ImageList = frmMaster.imlNode
    Set nodtemp = tvwSelection(intDimensions + 1).Nodes.Add(, , "root", "Column_" & intDimensions + 1, conDefaultImage)
    On Error Resume Next
    intDimensions = intDimensions + 1
    tvwSelection(intDimensions).ZOrder conZTop
    tabDimensions.Tabs(intDimensions).Selected = True
Exit Sub
'Error Handler
AddDimension_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub
Public Sub DeleteLastDimension()
'Created by Robin G Brown, 25/4/97
'Removes a tab and a dimension from the form
'Sub Declares
Const conSub = "DeleteLastDimension"
    'Error Trap
    On Error GoTo DeleteLastDimension_ErrorHandler
    If intDimensions = 1 Then
        MsgBox "Cannot delete last dimension."
        Exit Sub
    End If
    If tvwSelection(intDimensions).Nodes(1).Children > 0 Then
        If MsgBox("This dimension contains data, remove anyway?", vbOKCancel) = vbCancel Then
            Exit Sub
        End If
    End If
    tabDimensions.Tabs.Remove intDimensions
    Unload tvwSelection(intDimensions)
    On Error Resume Next
    intDimensions = intDimensions - 1
    tvwSelection(intDimensions).ZOrder conZTop
    tabDimensions.Tabs(intDimensions).Selected = True
Exit Sub
'Error Handler
DeleteLastDimension_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub
Public Sub DeleteDimension()
'Created by Robin G Brown, 25/4/97
'Removes a tab and a dimension from the form
'Sub Declares
Const conSub = "DeleteDimension"
Dim intDeleteDimension As Integer
    'Error Trap
    On Error GoTo DeleteDimension_ErrorHandler
    intDeleteDimension = tabDimensions.SelectedItem.Index
    If intDeleteDimension = 1 Then
        MsgBox "Cannot delete first dimension."
        Exit Sub
    End If
    'Need code here to sort out remaining dimensions - RGB/25/4/97
    tabDimensions.Tabs.Remove intDeleteDimension
    Unload tvwSelection(intDeleteDimension)
    On Error Resume Next
    intDimensions = intDimensions - 1
    tvwSelection(1).ZOrder conZTop
    tabDimensions.Tabs(1).Selected = True
Exit Sub
'Error Handler
DeleteDimension_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

