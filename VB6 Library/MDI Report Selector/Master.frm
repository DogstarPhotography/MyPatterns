VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.MDIForm frmMaster 
   BackColor       =   &H00808080&
   Caption         =   "Report Selector"
   ClientHeight    =   3045
   ClientLeft      =   1275
   ClientTop       =   1965
   ClientWidth     =   6510
   Icon            =   "Master.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrAuto 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1260
      Top             =   60
   End
   Begin ComctlLib.StatusBar stbStatus 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   2730
      Width           =   6510
      _ExtentX        =   11483
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327680
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   176
            TextSave        =   "11:22"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1376
            MinWidth        =   176
            TextSave        =   "20/05/97"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   13229
            MinWidth        =   13229
            Text            =   "Status"
            TextSave        =   "Status"
            Key             =   ""
            Object.Tag             =   ""
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
      MouseIcon       =   "Master.frx":0442
   End
   Begin ComctlLib.ImageList imlDefault 
      Left            =   60
      Top             =   660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Master.frx":045E
            Key             =   "rows"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Master.frx":0778
            Key             =   "cols"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Master.frx":0A92
            Key             =   "pages"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Master.frx":0DAC
            Key             =   "default"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Master.frx":10C6
            Key             =   "check"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Master.frx":13E0
            Key             =   "uncheck"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imlNode 
      Left            =   660
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327680
   End
   Begin ComctlLib.ImageList imlDimension 
      Left            =   60
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483638
      MaskColor       =   12632256
      _Version        =   327680
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "Close"
      End
      Begin VB.Menu mnuFileDatabase 
         Caption         =   "Database"
      End
      Begin VB.Menu mnuFileOpenSetting 
         Caption         =   "Use Setting"
         Begin VB.Menu mnuFileOpenSettingDefault 
            Caption         =   "Default"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuFileOpenSettingCustom 
            Caption         =   "Custom 1"
            Index           =   1
         End
         Begin VB.Menu mnuFileOpenSettingCustom 
            Caption         =   "Custom 2"
            Index           =   2
         End
         Begin VB.Menu mnuFileOpenSettingCustom 
            Caption         =   "Custom 3"
            Index           =   3
         End
      End
      Begin VB.Menu mnuFileSaveSetting 
         Caption         =   "Save Setting"
         Begin VB.Menu mnuFileSaveSettingCustom 
            Caption         =   "Custom 1"
            Index           =   1
         End
         Begin VB.Menu mnuFileSaveSettingCustom 
            Caption         =   "Custom 2"
            Index           =   2
         End
         Begin VB.Menu mnuFileSaveSettingCustom 
            Caption         =   "Custom 3"
            Index           =   3
         End
      End
      Begin VB.Menu mnuFileOptions 
         Caption         =   "&Options..."
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTree 
      Caption         =   "Tree"
      Begin VB.Menu mnuTreeExpand 
         Caption         =   "Expand Node"
      End
      Begin VB.Menu mnuTreeCollapse 
         Caption         =   "Collapse Node"
      End
      Begin VB.Menu mnuTreeDelete 
         Caption         =   "Delete Node"
      End
      Begin VB.Menu mnuTreeSelect 
         Caption         =   "Select Node"
      End
      Begin VB.Menu mnuTreeDeselect 
         Caption         =   "Deselect Node"
      End
      Begin VB.Menu mnuTreeDefault 
         Caption         =   "Set Default"
      End
   End
   Begin VB.Menu mnuDimension 
      Caption         =   "Dimension"
      Begin VB.Menu mnuDimensionAdd 
         Caption         =   "Add Dimension"
      End
      Begin VB.Menu mnuDimensionDelete 
         Caption         =   "Delete Dimension"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Window"
      Begin VB.Menu mnuWindowScratchpad 
         Caption         =   "ScratchPad"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuWindowAutoArrange 
         Caption         =   "Auto Arrange"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuWindowArrangeRoot 
         Caption         =   "Arrange"
         Begin VB.Menu mnuWindowArrange 
            Caption         =   "Cascade"
            Index           =   1
         End
         Begin VB.Menu mnuWindowArrange 
            Caption         =   "Tile Horizontal"
            Index           =   2
         End
         Begin VB.Menu mnuWindowArrange 
            Caption         =   "Tile Vertical"
            Index           =   3
         End
      End
      Begin VB.Menu mnuWindowSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowList 
         Caption         =   "Windows"
         WindowList      =   -1  'True
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "Report"
      Begin VB.Menu mnuReportFillGrid 
         Caption         =   "Fill Grid"
      End
      Begin VB.Menu mnuReportPrintGrid 
         Caption         =   "Print Grid"
      End
      Begin VB.Menu mnuReportReportOn 
         Caption         =   "Report on..."
      End
      Begin VB.Menu mnuReportTriReport 
         Caption         =   "TriReport"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "Contents"
      End
      Begin VB.Menu mnuHelpSearch 
         Caption         =   "Search for Help on..."
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Robin G Brown, 22nd April 1997
'-----------------------------------------------------------------------------
'   This form contains code for _
    the master MDI window
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "frmMaster"
Private intCounter As Integer
Private intReturn As Integer
Private booChildrenLoaded As Boolean
Private Sub MDIForm_Load()
'Created by Robin G Brown, 22/4/97
'Default behaviour
'Sub Declares
Const conSub = "Form_Load"
    'Error Trap
    On Error GoTo Form_Load_ErrorHandler
    booChildrenLoaded = False
    If strRegistrySettings(1) = "ON" Then
        mnuWindowAutoArrange.Checked = True
    Else
        mnuWindowAutoArrange.Checked = False
    End If
    If strRegistrySettings(2) = "ON" Then
        mnuWindowScratchpad.Checked = True
    Else
        mnuWindowScratchpad.Checked = False
    End If
Exit Sub
'Error Handler
Form_Load_ErrorHandler:
    Exit Sub
End Sub
Private Sub MDIForm_Resize()
'Created by Robin G Brown, 22/4/97
'Default behaviour
'Sub Declares
Const conSub = "MDIForm_Resize"
    'Error trap
    On Error GoTo MDIForm_Resize_ErrorHandler
    If Me.WindowState <> 1 Then
        If Me.Height < (Screen.Height / 2) Then Me.Height = Screen.Height / 2
        If Me.Width < (Screen.Width / 2) Then Me.Width = Screen.Width / 2
        If mnuWindowAutoArrange.Checked = True And booChildrenLoaded = True Then
            Call ArrangeChildren(0)
        End If
    End If
Exit Sub
MDIForm_Resize_ErrorHandler:
    Exit Sub
End Sub

Private Sub mnuDimensionAdd_Click()
'Created by Robin G Brown, 25/4/97
'Calls AddDimension for the relevant window
    'Error Trap
    On Error Resume Next
    If TypeOf Me.ActiveForm Is mdiSelectionChild Then
        Call ActiveForm.AddDimension
    End If
End Sub

Private Sub mnuDimensionDelete_Click()
'Created by Robin G Brown, 25/4/97
'Calls DeleteLastDimension for the relevant window
    'Error Trap
    On Error Resume Next
    If TypeOf Me.ActiveForm Is mdiSelectionChild Then
        Call ActiveForm.DeleteLastDimension
        'Call ActiveForm.DeleteDimension
    End If
End Sub

Private Sub mnuFileClose_Click()
'Created by Robin G Brown, 28/4/97
'Get rid of the current crop of windows
    'Error Trap
    On Error Resume Next
    'Save any settings?
    'CODE_HERE
    'Unload those pesky kids
    Call UnloadChildren
End Sub

Private Sub mnuFileDatabase_Click()
'Created by Robin G Brown, 28/4/97
'Database functions are not available
    'Error Trap
    On Error Resume Next
    'CODE_HERE
    Call Unavailable
End Sub

Private Sub mnuFileExit_Click()
'Created by Robin G Brown, 25/4/97
'End the application
    'Error Trap
    On Error Resume Next
    Call ApplicationEnd
End Sub

Public Sub mnuFileNew_Click()
'Created by Robin G Brown, 28/4/97
'Display a new set of windows
    'Error Trap
    On Error Resume Next
    Call UnloadChildren
    Call LoadChildren
    Call ArrangeChildren(0)
End Sub

Private Sub mnuFileOpenSettingCustom_Click(Index As Integer)
'Created by Robin G Brown, 25/4/97
'Arrange the windows according to the Custom setting
    'Error Trap
    On Error Resume Next
    Call ArrangeChildren(Index)
End Sub

Private Sub mnuFileOpenSettingDefault_Click()
'Created by Robin G Brown, 25/4/97
'Arrange the windows according to the default setting
    'Error Trap
    On Error Resume Next
    Call ArrangeChildren(0)
End Sub

Private Sub mnuFileOptions_Click()
'Created by Robin G Brown, 28/4/97
'Show the options form
    'Error Trap
    On Error Resume Next
    frmOptions.Show 1
End Sub

Private Sub mnuFileSaveSettingCustom_Click(Index As Integer)
'Created by Robin G Brown, 25/4/97
'Save the current window arrangement in the registry, under the given heading
    'Error Trap
    On Error Resume Next
    MsgBox "Unable to save settings for Custom " & Index
End Sub

Private Sub mnuFileStyle_Click(Index As Integer)
'Created by Robin G Brown, 25/4/97
'Change all the pictures around according to the style chosen
    'Error Trap
    On Error Resume Next
    Call ChangeTheme(Index)
End Sub

Private Sub mnuHelpAbout_Click()
'Created by Robin G Brown, 25/4/97
'Display the about form
    'Error Trap
    On Error Resume Next
    frmAbout.Show vbModal
End Sub

Private Sub mnuHelpContents_Click()
    Call Unavailable
End Sub

Private Sub mnuHelpSearch_Click()
    Call Unavailable
End Sub

Private Sub mnuReportFillGrid_Click()
'Created by Robin G Brown, 25/4/97
'Calls fillreportgrid
    'Error Trap
    On Error Resume Next
    Call mdiReportGrid.FillReportGrid
End Sub

Private Sub mnuReportPrintGrid_Click()
    Call Unavailable
End Sub

Private Sub mnuReportReportOn_Click()
'Created by Robin G Brown, 25/4/97
'Message for relevant report
Dim strReportTitle As String
    'Error Trap
    On Error Resume Next
    strReportTitle = mdiReportGrid.GetReportTitle()
    If strReportTitle <> "" Then
        intReturn = MsgBox("Report on " & vbLf & strReportTitle, vbYesNo)
        If intReturn = vbYes Then
            Call ExcelReport(strReportTitle)
        End If
    Else
        MsgBox "Cannot retrieve report."
    End If
End Sub

Private Sub mnuReportTriReport_Click()
    Call ShowTriangulator
End Sub

Private Sub mnuTreeCollapse_Click()
'Created by Robin G Brown, 25/4/97
'Collapse the tree up to the current node
    'Error Trap
    On Error Resume Next
    If tvwCurrent.SelectedItem.Key <> "" Then
        Call CollapseBranch(tvwCurrent, tvwCurrent.SelectedItem.Key)
    End If
End Sub

Private Sub mnuTreeDefault_Click()
'Created by Robin G Brown, 25/4/97
'Set the image for the selected node
    'Error Trap
    On Error Resume Next
    If nodDrag.Selected = True Then
        nodDrag.Image = conDefault
    End If
End Sub

Private Sub mnuTreeDelete_Click()
'Created by Robin G Brown, 25/4/97
'Delete the selected node
    'Error Trap
    On Error Resume Next
    'Delete the node unless it is in the source control
    If tvwCurrent Is mdiTreeChild.tvwSource Then
        Beep
    Else
        tvwCurrent.Nodes.Remove nodDrag.Key
        Set tvwCurrent = Nothing
        Set nodDrag = Nothing
    End If
End Sub

Private Sub mnuTreeDeselect_Click()
'Created by Robin G Brown, 25/4/97
'Set the image for the selected node
    'Error Trap
    On Error Resume Next
    If nodDrag.Selected = True Then
        nodDrag.Image = conDeselected
    End If
End Sub

Private Sub mnuTreeExpand_Click()
'Created by Robin G Brown, 25/4/97
'Expand the tree from the current node
    'Error Trap
    On Error Resume Next
    If tvwCurrent.SelectedItem.Key <> "" Then
        Call ExpandBranch(tvwCurrent, tvwCurrent.SelectedItem.Key)
    End If
End Sub

Private Sub mnuTreeSelect_Click()
'Created by Robin G Brown, 25/4/97
'Set the image for the selected node
    'Error Trap
    On Error Resume Next
    If nodDrag.Selected = True Then
        nodDrag.Image = conSelected
    End If
End Sub

Private Sub mnuWindowArrange_Click(Index As Integer)
'Created by Robin G Brown, 25/4/97
'Arrange the windows in the tile style
    'Error Trap
    On Error Resume Next
    Select Case Index
    Case 1
        Call ArrangeChildren(11)
    Case 2
        Call ArrangeChildren(12)
    Case 3
        Call ArrangeChildren(13)
    Case Else
        Call ArrangeChildren(11)
    End Select
End Sub

Private Sub mnuWindowAutoArrange_Click()
'Created by Robin G Brown, 25/4/97
'Set the checkmark on or off
    'Error Trap
    On Error Resume Next
    If mnuWindowAutoArrange.Checked = True Then
        mnuWindowAutoArrange.Checked = False
        strRegistrySettings(1) = "OFF"
        tmrAuto.Enabled = False
    Else
        mnuWindowAutoArrange.Checked = True
        strRegistrySettings(1) = "ON"
        'Arrange the windows
        Call ArrangeChildren(0)
        tmrAuto.Enabled = True
    End If
End Sub

Private Sub mnuWindowScratchpad_Click()
'Created by Robin G Brown, 25/4/97
'Set the checkmark on or off
    'Error Trap
    On Error Resume Next
    If mnuWindowScratchpad.Checked = True Then
        mnuWindowScratchpad.Checked = False
        strRegistrySettings(2) = "OFF"
    Else
        mnuWindowScratchpad.Checked = True
        strRegistrySettings(2) = "ON"
    End If
    'Resize the scratchpad
    Call mdiTreeChild.ArrangeSourceAndScratch
End Sub
Public Sub ArrangeChildren(Optional ByVal varSetting As Variant)
'Created by Robin G Brown, 23/4/97
'Arranges child windows according to settings in registry
'Constants defined in modMain:
'Public Const conAutoArrange = 99
'Public Const conDefault = 0
'Public Const conCustom1 = 1
'Public Const conCustom2 = 2
'Public Const conCustom3 = 3
'Public Const conCascade = 11
'Public Const conTileHorizontal = 12
'Public Const conTileVertical = 13
'Dim frmPosition(1 To 6) As Form
'Spoiler info for auto arrange...
'
'   ---------------------------------   <- 0
'   |       |   2   |  4    |       |
'   |       |       |       |       |
'   |   1   |-------|-------|  6    |   <- intVertSplit
'   |       |   3   |  5    |       |
'   |       |       |       |       |
'   ---------------------------------   <- intMaxHeight
'   ^               ^               ^
'   0               intHorSplit     intmaxwidth
'           ^               ^
'           intLeftSplit    intRightOffset
'
'End of spoiler info.
'Sub Declares
Dim intMaxHeight As Integer
Dim intMaxWidth As Integer
Dim intSetting As Integer
Dim intVertSplit As Integer
Dim intHorSplit As Integer
Dim intLeftsplit As Integer
Dim intRightOffset As Integer
    'Error Trap
    On Error Resume Next
    If IsMissing(varSetting) Then
        intSetting = 99
    Else
        If VarType(varSetting) = vbInteger Then
            intSetting = varSetting
        Else
            intSetting = 99
        End If
    End If
    Select Case intSetting
    Case 0
        'Default setting
        Set frmPosition(1) = mdiTreeChild
        Set frmPosition(2) = mdiPageSelector
        Set frmPosition(3) = mdiRowSelector
        Set frmPosition(4) = mdiColSelector
        Set frmPosition(5) = mdiReportGrid
        Set frmPosition(6) = Nothing
        'Auto arrange
        intMaxHeight = Me.Height - 1080
        intMaxWidth = Me.Width - 180
        'Only resize if all of the forms are 'normal'
        'frmPosition(1) is either on left or right, which determines where all other windows go...
        intVertSplit = frmPosition(2).Height
        intHorSplit = frmPosition(2).Width
        'frmPosition(1) on left
        intLeftsplit = frmPosition(1).Width
        intRightOffset = intMaxWidth
        frmPosition(1).Move _
            0, _
            0, _
            frmPosition(1).Width, _
            intMaxHeight
        frmPosition(2).Move _
            intLeftsplit, _
            0, _
            intHorSplit, _
            intVertSplit
        frmPosition(3).Move _
            intLeftsplit, _
            intVertSplit, _
            intHorSplit, _
            intMaxHeight - intVertSplit
        frmPosition(4).Move _
            intLeftsplit + intHorSplit, _
            0, _
            intRightOffset - (intLeftsplit + intHorSplit), _
            intVertSplit
        frmPosition(5).Move _
            intLeftsplit + intHorSplit, _
            intVertSplit, _
            intRightOffset - (intLeftsplit + intHorSplit), _
            intMaxHeight - intVertSplit
    Case 1
        'Retrieve settings
        'Move/size windows
        MsgBox "Custom setting 1 not available."
    Case 2
        'Retrieve settings
        'Move/size windows
        MsgBox "Custom setting 2 not available."
    Case 3
        'Retrieve settings
        'Move/size windows
        MsgBox "Custom setting 3 not available."
    Case 11
        Me.Arrange vbCascade
    Case 12
        Me.Arrange vbTileHorizontal
    Case 13
        Me.Arrange vbTileVertical
    Case Else 'also 99 = conAutoArrange
        'Auto arrange
        intMaxHeight = Me.Height - 1080
        intMaxWidth = Me.Width - 180
        'Only resize if all of the forms are 'normal'
        'frmPosition(1) is either on left or right, which determines where all other windows go...
        'frmPosition(2) determines where some windows go so prevent it becoming too large
        intVertSplit = frmPosition(2).Height
        'If intVertSplit > Me.Height * 0.8 Then
        '
        'End If
        intHorSplit = frmPosition(2).Width
        If Not frmPosition(1) Is Nothing Then
            'frmPosition(1) on left
            intLeftsplit = frmPosition(1).Width
            intRightOffset = intMaxWidth
            frmPosition(1).Move _
                0, _
                0, _
                frmPosition(1).Width, _
                intMaxHeight
            frmPosition(2).Move _
                intLeftsplit, _
                0, _
                intHorSplit, _
                intVertSplit
            frmPosition(3).Move _
                intLeftsplit, _
                intVertSplit, _
                intHorSplit, _
                intMaxHeight - intVertSplit
            frmPosition(4).Move _
                intLeftsplit + intHorSplit, _
                0, _
                intRightOffset - (intLeftsplit + intHorSplit), _
                intVertSplit
            frmPosition(5).Move _
                intLeftsplit + intHorSplit, _
                intVertSplit, _
                intRightOffset - (intLeftsplit + intHorSplit), _
                intMaxHeight - intVertSplit
        Else
            'frmPosition(6) on right
            intLeftsplit = 0
            intRightOffset = intMaxWidth - frmPosition(6).Width
            frmPosition(6).Move _
                intRightOffset, _
                0, _
                frmPosition(6).Width, _
                intMaxHeight
            frmPosition(2).Move _
                intLeftsplit, _
                0, _
                intHorSplit, _
                intVertSplit
            frmPosition(3).Move _
                intLeftsplit, _
                intVertSplit, _
                intHorSplit, _
                intMaxHeight - intVertSplit
            frmPosition(4).Move _
                intLeftsplit + intHorSplit, _
                0, _
                intRightOffset - (intLeftsplit + intHorSplit), _
                intVertSplit
            frmPosition(5).Move _
                intLeftsplit + intHorSplit, _
                intVertSplit, _
                intRightOffset - (intLeftsplit + intHorSplit), _
                intMaxHeight - intVertSplit
        End If
    End Select
End Sub
Public Sub BackupOfArrangeChildren(Optional ByVal varSetting As Variant)
'Created by Robin G Brown, 23/4/97
'Arranges child windows according to settings in registry
'Constants defined in modMain:
'Public Const conAutoArrange = 99
'Public Const conDefault = 0
'Public Const conCustom1 = 1
'Public Const conCustom2 = 2
'Public Const conCustom3 = 3
'Public Const conCascade = 11
'Public Const conTileHorizontal = 12
'Public Const conTileVertical = 13
'Dim frmPosition(1 To 6) As Form
'Spoiler info for auto arrange...
'
'   ---------------------------------   <- 0
'   |       |   2   |  3    |       |
'   |       |       |       |       |
'   |   1   |-------|-------|  6    |   <- intVertSplit
'   |       |   4   |  5    |       |
'   |       |       |       |       |
'   ---------------------------------   <- intMaxHeight
'   ^               ^               ^
'   0               intHorSplit     intmaxwidth
'           ^               ^
'           intLeftSplit    intRightOffset
'
'End of spoiler info.
'Sub Declares
Dim intMaxHeight As Integer
Dim intMaxWidth As Integer
Dim intSetting As Integer
Dim intVertSplit As Integer
Dim intHorSplit As Integer
Dim intLeftsplit As Integer
Dim intRightOffset As Integer
    'Error Trap
    On Error Resume Next
    If IsMissing(varSetting) Then
        intSetting = 99
    Else
        If VarType(varSetting) = vbInteger Then
            intSetting = varSetting
        Else
            intSetting = 99
        End If
    End If
    Select Case intSetting
    Case 1
        'Retrieve settings
        'Move/size windows
        MsgBox "Custom setting 1 not available."
    Case 2
        'Retrieve settings
        'Move/size windows
        MsgBox "Custom setting 2 not available."
    Case 3
        'Retrieve settings
        'Move/size windows
        MsgBox "Custom setting 3 not available."
    Case 11
        Me.Arrange vbCascade
    Case 12
        Me.Arrange vbTileHorizontal
    Case 13
        Me.Arrange vbTileVertical
    Case Else 'also 99 or 0
        'Use default
        intMaxHeight = Me.Height - 1080
        intMaxWidth = Me.Width - 180
        If mdiTreeChild.WindowState = 0 _
        And mdiPageSelector.WindowState = 0 _
        And mdiRowSelector.WindowState = 0 _
        And mdiColSelector.WindowState = 0 _
        And mdiReportGrid.WindowState = 0 _
        Then
            'Only resize if all of the forms are 'normal'
            'mdiTreeChild is either on left or right, which determines where all other windows go...
            intVertSplit = mdiPageSelector.Height
            intHorSplit = mdiPageSelector.Width
            If mdiTreeChild.Left < (intMaxWidth / 2) Then
                'mdiTreeChild on left
                intLeftsplit = mdiTreeChild.Width
                intRightOffset = intMaxWidth
                mdiTreeChild.Move _
                    0, _
                    0, _
                    mdiTreeChild.Width, _
                    intMaxHeight
            Else
                'mdiTreeChild on right
                intLeftsplit = 0
                intRightOffset = intMaxWidth - mdiTreeChild.Width
                mdiTreeChild.Move _
                    intRightOffset, _
                    0, _
                    mdiTreeChild.Width, _
                    intMaxHeight
            End If
            mdiPageSelector.Move _
                intLeftsplit, _
                0, _
                intHorSplit, _
                intVertSplit
            mdiRowSelector.Move _
                intLeftsplit, _
                intVertSplit, _
                intHorSplit, _
                intMaxHeight - intVertSplit
            mdiColSelector.Move _
                intLeftsplit + intHorSplit, _
                0, _
                intRightOffset - (intLeftsplit + intHorSplit), _
                intVertSplit
            mdiReportGrid.Move _
                intLeftsplit + intHorSplit, _
                intVertSplit, _
                intRightOffset - (intLeftsplit + intHorSplit), _
                intMaxHeight - intVertSplit
        End If
    End Select
End Sub
Public Sub LoadChildren()
'Created by Robin G Brown, 28/4/97
'Display a set of windows
'Sub Declares
Const conSub = "LoadChildren"
    'Error Trap
    On Error GoTo LoadChildren_ErrorHandler
    'Retrieve any settings
    'CODE_HERE
    'Load those pesky kids
    booChildrenLoaded = False
    Set mdiRowSelector = New mdiSelectionChild
    Set mdiColSelector = New mdiSelectionChild
    Set mdiPageSelector = New mdiSelectionChild
    mdiTreeChild.Show
    mdiRowSelector.Show
    mdiColSelector.Show
    mdiPageSelector.Show
    mdiReportGrid.Show
    booChildrenLoaded = True
    If mnuWindowAutoArrange.Checked = False Then
        tmrAuto.Enabled = False
    Else
        'Arrange the windows
        Call MDIForm_Resize
        tmrAuto.Enabled = True
    End If
Exit Sub
'Error Handler
LoadChildren_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub
Public Sub UnloadChildren()
'Created by Robin G Brown, 28/4/97
'Get rid of the current crop of windows
'Sub Declares
Const conSub = "UnloadChildren"
    'Error Trap
    On Error GoTo UnloadChildren_ErrorHandler
    'Save any settings?
    tmrAuto.Enabled = False
    'CODE_HERE
    'Unload those pesky kids
    Unload mdiReportGrid
    Unload mdiPageSelector
    Unload mdiColSelector
    Unload mdiRowSelector
    Unload mdiTreeChild
    booChildrenLoaded = False
Exit Sub
'Error Handler
UnloadChildren_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Private Sub tmrAuto_Timer()
'Created by Robin G Brown, 28/4/97
'This timer is used to 'lock' the forms in position by rearranging them each tinme this event occurs
'Sub Declares
    'Error Trap
    On Error Resume Next
    Call ArrangeChildren(conAutoArrange)
End Sub
