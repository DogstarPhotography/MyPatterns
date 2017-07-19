VERSION 4.00
Begin VB.Form frmSelector 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Selector"
   ClientHeight    =   4665
   ClientLeft      =   7620
   ClientTop       =   1755
   ClientWidth     =   3960
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Height          =   5355
   Left            =   7560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4665
   ScaleWidth      =   3960
   Top             =   1125
   Width           =   4080
   Begin ComctlLib.TreeView tvwDimension 
      DragIcon        =   "SelectorChild.frx":0000
      Height          =   3255
      Index           =   1
      Left            =   60
      TabIndex        =   1
      Top             =   360
      Width           =   3795
      _Version        =   65536
      _ExtentX        =   6694
      _ExtentY        =   5741
      _StockProps     =   196
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      HideSelection   =   0   'False
      ImageList       =   ""
      Indentation     =   176
      LabelEdit       =   1
      PathSeparator   =   "\"
      Style           =   7
   End
   Begin ComctlLib.TabStrip tabDimensions 
      Height          =   3675
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3915
      _Version        =   65536
      _ExtentX        =   6906
      _ExtentY        =   6482
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImageList       =   ""
      ShowTips        =   0   'False
      i1              =   "SelectorChild.frx":0442
   End
   Begin VB.Menu mnuChildFile 
      Caption         =   "File"
   End
End
Attribute VB_Name = "frmSelector"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'Created by Robin G Brown, 22nd April 1997
'-----------------------------------------------------------------------------
'   This form contains code for _
    child form for manipulating a tree control
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "frmTreeChild"
Private intCounter As Integer
Private intReturn As Integer
Private intDimensions As Integer
Private intDimensionImage As Integer
Private Sub Form_Load()
'Created by Robin G Brown, 22/4/97
'Populate the treeview controls
'Sub Declares
Const conSub = "Form_Load"
Dim nodTemp As Node
    'Error Trap
    On Error GoTo Form_Load_ErrorHandler
    'The only difference between forms is the icon on the tabs!
    Select Case Forms.Count - 3
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
    tabDimensions.ImageList = GetDimensionThemeImageList()
    tabDimensions.Tabs(1).Image = intDimensionImage
    intDimensions = 3
    For intCounter = 1 To intDimensions
        If intCounter > 1 Then
            tabDimensions.Tabs.Add intCounter, "col" & intCounter, CStr(intCounter), intDimensionImage
            'tabDimensions.ZOrder conZBottom
            Load tvwDimension(intCounter)
        End If
        tvwDimension(intCounter).Visible = True
        'tvwDimension(intCounter).Visible = False
        'tvwDimension(intCounter).ZOrder conZTop
        tvwDimension(intCounter).ImageList = GetTreeThemeImageList()
        Set nodTemp = tvwDimension(intCounter).Nodes.Add(, , "root", "Column_" & intCounter, conDefaultImage)
    Next
    'tvwDimension(1).Visible = True
    tvwDimension(1).ZOrder conZTop
    Me.Height = 3600
    Me.Width = 3600
Exit Sub
'Error Handler
Form_Load_ErrorHandler:
    MsgBox "Unexpected error : " & Error$(Err), vbExclamation, conModule & " : " & conSub
    Exit Sub
End Sub
Private Sub Form_Resize()
'Created by Robin G Brown, 22/4/97
'Resize controls on form
'Sub Declares
Const conSub = "Form_Resize"
    'Error Trap
    On Error GoTo Form_Resize_ErrorHandler
    If Me.Height < 1000 Then Me.Height = 1000
    If Me.Width < 1000 Then Me.Width = 1000
    tabDimensions.Height = Me.Height - 420
    tabDimensions.Width = Me.Width - 120
    For intCounter = 1 To intDimensions
        tvwDimension(intCounter).Height = Me.Height - 820
        tvwDimension(intCounter).Width = Me.Width - 240
    Next
Exit Sub
'Error Handler
Form_Resize_ErrorHandler:
    MsgBox "Unexpected error : " & Error$(Err), vbExclamation, conModule & " : " & conSub
    Exit Sub
End Sub
Private Sub tabDimensions_Click()
'Created by Robin G Brown, 22/4/97
'Handle someone clicking on the tab, move treeview controls in ZOrder as appropriate
'Sub Declares
Const conSub = "ImmediateErrorSub"
    'Error Trap
    On Error GoTo ImmediateErrorSub_ErrorHandler
    'Indices on the tabstrip control match those for the treeview controls
    'For intCounter = 1 To intDimensions
    '    'tabDimensions.ZOrder conZBottom
    '    If intCounter = tabDimensions.SelectedItem.Index Then
    '        'tvwDimension(intCounter).Visible = True
    '        tvwDimension(intCounter).ZOrder conZTop
    '    Else
    '        'tvwDimension(intCounter).Visible = False
    '        tvwDimension(intCounter).ZOrder conZBottom
    '    End If
    'Next
    tvwDimension(tabDimensions.SelectedItem.Index).ZOrder conZTop
    'Me.Refresh
Exit Sub
'Error Handler
ImmediateErrorSub_ErrorHandler:
    Exit Sub
End Sub
Private Sub tvwDimension_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
'Created by Robin G Brown, 15/4/97
'Handles moving nodes around using dragdrop functionality
'Sub Declares
Const conSub = "tvwDimension_DragDrop"
    On Error GoTo tvwDimension_DragDrop_ErrorHandler
    'Drop the source onto the target
    If Not tvwDimension(Index).DropHighlight Is Nothing Then
        'If the target is not the source
        If Not nodDrag = tvwDimension(Index).DropHighlight Then
            If tvwDimension(Index) Is Source Then
                'If we are moving within the same control, just change the parent
                Set nodDrag.Parent = tvwDimension(Index).DropHighlight
            Else
                'Validate the action, as some movements are not allowed, _
                    assume that we are dropping onto one of the 'target' controls
                'A series of validation questions, which must all be 'passed' in order to complete the operation
                'VALIDATION_CODE_HERE
                'Copy whole branch to target and delete source node/branch, unless it came from tvwDimension(0)
                Call CopyBranch(tvwDimension(Index), nodDrag, tvwDimension(Index).DropHighlight.Key)
                'Delete source branch if it is not from tvwSource
                If Not Source Is frmTreeChild.tvwSource Then
                    Source.Nodes.Remove nodDrag.Key
                End If
            End If
        End If
    End If
    'Finish off but don't throw an error
    On Error Resume Next
    'Show the dropped nodes in some way
    If nodDrag.Expanded = True Then
        'Call ExpandBranch(tvwDimension(Index), tvwDimension(Index).DropHighlight.Key)
        tvwDimension(Index).DropHighlight.Child.EnsureVisible
    Else
        tvwDimension(Index).DropHighlight.Child.EnsureVisible
    End If
    Set tvwDimension(Index).DropHighlight = Nothing
    Set Source.DropHighlight = Nothing
    tvwDimension(Index).Refresh
    Source.Refresh
    booDrag = False
    StatusBar ""
Exit Sub
tvwDimension_DragDrop_ErrorHandler:
    If Err = 55555 Then
        MsgBox "That action is not allowed."
    Else
        MsgBox "Invalid operation, please reset"
    End If
    On Error Resume Next
    Set tvwDimension(Index).DropHighlight = Nothing
    Set Source.DropHighlight = Nothing
    tvwDimension(Index).Refresh
    Source.Refresh
    booDrag = False
    Exit Sub
End Sub
Private Sub tvwDimension_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)
    'Set index to top control, to force correct selection of controls!
    Index = tabDimensions.SelectedItem.Index
    If booDrag = True Then
        'Set DropHighlight to the mouse's coordinates
        Set tvwDimension(Index).DropHighlight = tvwDimension(Index).HitTest(x, y)
        If Not tvwDimension(Index).DropHighlight Is Nothing Then
            StatusBar "Release the mouse button to drop onto : " & tvwDimension(Index).DropHighlight.Text
        End If
    End If
End Sub
Private Sub tvwDimension_GotFocus(Index As Integer)
    'Set index to top control, to force correct selection of controls!
    Index = tabDimensions.SelectedItem.Index
    Set tvwCurrent = tvwDimension(Index)
End Sub
Private Sub tvwDimension_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    'Set index to top control, to force correct selection of controls!
    Index = tabDimensions.SelectedItem.Index
    'Select the node
    Set nodDrag = tvwDimension(Index).HitTest(x, y)
    Set tvwCurrent = tvwDimension(Index)
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
Private Sub tvwDimension_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    'Set index to top control, to force correct selection of controls!
    Index = tabDimensions.SelectedItem.Index
    'Start the dragdrop operation
    If Button = vbLeftButton Then
        If nodDrag.Parent Is Nothing Then
            'Do not allow drag operation, i.e. do nothing
        Else
            booDrag = True
            tvwDimension(Index).Drag vbBeginDrag
        End If
    End If
End Sub


