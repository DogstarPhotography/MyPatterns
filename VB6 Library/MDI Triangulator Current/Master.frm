VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Begin VB.MDIForm frmMaster 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Triangulator"
   ClientHeight    =   6570
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10005
   Icon            =   "Master.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar tlbMaster 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlSmallToolbar"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   3
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   1200
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "otheryears"
            Object.ToolTipText     =   "Other Years"
            Object.Tag             =   ""
            ImageIndex      =   2
            Style           =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   1695
            MixedState      =   -1  'True
         EndProperty
      EndProperty
      Begin VB.ComboBox cmbChartType 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   60
         Width           =   1695
      End
      Begin VB.ComboBox cmbStartYear 
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   60
         Width           =   1095
      End
   End
   Begin MSComDlg.CommonDialog dlgMaster 
      Left            =   60
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin ComctlLib.StatusBar stbStatus 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   6255
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   12012
            Key             =   "panStatus"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "No Other Years..."
            TextSave        =   "No Other Years..."
            Key             =   "panYears"
            Object.Tag             =   "OFF"
            Object.ToolTipText     =   "Click here to change multiple year selection"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "18/09/97"
            Key             =   "panDateTime"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imlSmallToolbar 
      Left            =   60
      Top             =   1620
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Master.frx":0442
            Key             =   "cube"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Master.frx":075C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Master.frx":0A76
            Key             =   "graph"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Master.frx":0D90
            Key             =   "grid"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imlToolbar 
      Left            =   60
      Top             =   1020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Master.frx":10AA
            Key             =   "cube"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Master.frx":13C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Master.frx":16DE
            Key             =   "graph"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Master.frx":19F8
            Key             =   "grid"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFileRoot 
      Caption         =   "&File"
      Begin VB.Menu mnufile 
         Caption         =   "&New Report"
         Index           =   1
      End
      Begin VB.Menu mnufile 
         Caption         =   "New &Graph"
         Index           =   2
      End
      Begin VB.Menu mnufile 
         Caption         =   "&Open..."
         Index           =   3
      End
      Begin VB.Menu mnufile 
         Caption         =   "&Save..."
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu mnufile 
         Caption         =   "Save Session"
         Enabled         =   0   'False
         Index           =   5
      End
      Begin VB.Menu mnufile 
         Caption         =   "&Close"
         Index           =   6
      End
      Begin VB.Menu mnufile 
         Caption         =   "E&xit"
         Index           =   7
      End
   End
   Begin VB.Menu mnuWindowRoot 
      Caption         =   "&Window"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuHelpRoot 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Robin G Brown, 12th May 1997
'-----------------------------------------------------------------------------
'   This form contains code for _
    handling the various child forms of the triangulator app
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "frmMaster"
Private intCounter As Integer
Private intReturn As Integer
'Hold a reference for saving the session
Private sffSession As SecFile

Private Sub cmbChartType_Click()
'Created by Robin G Brown, 27/5/97
'Change the graph style
'Sub Declares
Const conSub = "cmbChartType_Click"
    'Error Trap
    On Error Resume Next
    If ActiveForm Is Nothing Then Exit Sub
    If TypeOf ActiveForm Is frmChartReport Then
        Select Case cmbChartType.ListIndex
        Case 0
            '2D Line
            ActiveForm.graReport.chartType = VtChChartType2dLine
        Case 1
            '3D Line
            ActiveForm.graReport.chartType = VtChChartType3dLine
        Case 2
            '3D step
            ActiveForm.graReport.chartType = VtChChartType3dStep
        Case 3
            '2D Step
            ActiveForm.graReport.chartType = VtChChartType2dStep
        Case 4
            '3D Area
            ActiveForm.graReport.chartType = VtChChartType3dArea
        Case 5
            '2D Area
            ActiveForm.graReport.chartType = VtChChartType2dArea
        Case 6
            '3D Bar
            ActiveForm.graReport.chartType = VtChChartType3dBar
        Case 7
            '2D Bar
            ActiveForm.graReport.chartType = VtChChartType2dBar
        Case Else
            '2D Line
            ActiveForm.graReport.chartType = VtChChartType2dLine
        End Select
        ActiveForm.ColouriseChart
        ActiveForm.graReport.Refresh
    End If
End Sub

Private Sub cmbStartYear_Click()
'Created by Robin G Brown, 27/5/97
'Update the current grid with the selected start year
'Sub Declares
Const conSub = "cmbStartYear_Click"
    'Error Trap
    On Error Resume Next
    If TypeOf ActiveForm Is frmGridReport Then
        Call ActiveForm.DataRandomInitialize(frmMaster.cmbStartYear.Text)
    End If
End Sub

Private Sub MDIForm_DragDrop(Source As Control, X As Single, Y As Single)
    'Create a chart?
End Sub

Private Sub MDIForm_Load()
'Created by Robin G Brown, 12/5/97
'Default Behaviour
'Sub Declares
Const conSub = "Form_Load"
Dim intCurrentYear As Integer
    'Error Trap
    On Error GoTo Form_Load_ErrorHandler
    'A couple of minor setups
    Call SetStatus("Ready")
    With stbStatus.Panels("panYears")
        .Tag = "OFF"
        .Text = "No Other Years"
        .ToolTipText = "" '"No Other Years"
    End With
    tlbMaster.Buttons("otheryears").Value = tbrUnpressed
    'Set up the toolbar and its elements
    'Load cmbStartYear with some years
    intCurrentYear = CInt(Format$(Now, "yyyy"))
    For intCounter = intCurrentYear To intCurrentYear - 15 Step -1
        cmbStartYear.AddItem (intCounter)
    Next
    cmbStartYear.ListIndex = 9
    'Load cmbChartType with some graph styles
    cmbChartType.AddItem "2D Line"
    cmbChartType.AddItem "3D Line"
    cmbChartType.AddItem "3D Step"
    cmbChartType.AddItem "2D Step"
    cmbChartType.AddItem "3D Area"
    cmbChartType.AddItem "2D Area"
    cmbChartType.AddItem "3D Bar"
    cmbChartType.AddItem "2D Bar"
    cmbChartType.ListIndex = 0
    cmbChartType.Enabled = False
Exit Sub
'Error Handler
Form_Load_ErrorHandler:
    Exit Sub
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
'Created by Robin G Brown, 12/5/97
'Default Behaviour
'Function Declares
Const conFunction = "Form_Unload"
Dim lngTempValue As Long
    'Errors 'disabled' to ensure that form unloads
    On Error Resume Next
    Set sffSession = Nothing
    Call ApplicationEnd
End Sub

Private Sub mnuFileRoot_Click()
'Created by Robin G Brown, 19/8/97
'Disable the Display Accounts item if a grid form is not showing
    'Error Trap
    On Error Resume Next
    If FormInstances("frmChartReport") > 0 Then
        mnuFile(2).Enabled = False
    Else
        mnuFile(2).Enabled = True
    End If
End Sub

Private Sub mnuFile_Click(Index As Integer)
'Created by Robin G Brown, 28/8/97
'Call appropriate menu routines in modCommonMenus module
'Sub Declares
Const conSub = "mnuFile_Click"
    'Error Trap
    On Error GoTo mnuFile_Click_ErrorHandler
    'Call the selected routine passing a reference to the current form
    Select Case Index
    Case 1
        Call modCommonMenus.NewReportMenuRoutine
    Case 2
        Call modCommonMenus.NewGraphMenuRoutine
    Case 3
        Call modCommonMenus.FileOpenMenuRoutine
    Case 6
        Unload frmMaster.ActiveForm
    Case 7
        Unload frmMaster
    Case Else
        'Default - do nothing
    End Select
Exit Sub
mnuFile_Click_ErrorHandler:
    MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    Exit Sub
End Sub

Private Sub mnuHelpAbout_Click()
'Created by Robin G Brown, 15/9/97
'Show the about form
    'Error Trap
    On Error Resume Next
    frmAbout.Show
End Sub

Private Sub stbStatus_PanelClick(ByVal Panel As ComctlLib.Panel)
'Created by Robin G Brown, 14/5/97
'Switch AndOtherYears on/off
'Sub Declares
    'Error Trap
    On Error Resume Next
    With Panel
        Select Case .Key
        Case "panStatus"
            '.Tag = ""
            .Text = ""
            '.ToolTipText = ""
        Case "panDateTime"
            If .Style = sbrDate Then
                .Style = sbrTime
            Else
                .Style = sbrDate
            End If
        Case "panYears"
            If .Tag = "ON" Then
                .Tag = "OFF"
                .Text = "No Other Years"
                tlbMaster.Buttons.Item(2).Value = tbrUnpressed
                '.ToolTipText = "No Other Years"
            Else
                .Tag = "ON"
                .Text = "And Other Years"
                tlbMaster.Buttons.Item(2).Value = tbrPressed
                '.ToolTipText = "And Other Years"
            End If
        Case Else
            'Do nothing
        End Select
    End With
End Sub

Public Sub SetStatus(ByVal strStatus As String)
'Created by Robin G Brown, 14/5/97
'Set the contants of the status bar
'Sub Declares
    'Error Trap
    On Error Resume Next
    With stbStatus.Panels("panStatus")
        '.Tag = strStatus
        .Text = strStatus
        '.ToolTipText = strStatus
    End With
End Sub

Private Sub tlbMaster_ButtonClick(ByVal Button As ComctlLib.Button)
'Created by Robin G Brown, 27/5/97
'Handle various buttons being clicked
'Sub Declares
Const conSub = "tlbMaster_ButtonClick"
    'Error Trap
    On Error GoTo tlbMaster_ButtonClick_ErrorHandler
    Select Case Button.Key
    Case "graph"
        'Do nothing
    Case "otheryears"
        With stbStatus.Panels("panYears")
            If .Tag = "ON" Then
                .Tag = "OFF"
                .Text = "No Other Years"
                '.ToolTipText = "No Other Years"
                Button.Value = tbrUnpressed
            Else
                .Tag = "ON"
                .Text = "And Other Years"
                '.ToolTipText = "And Other Years"
                Button.Value = tbrPressed
            End If
        End With
    Case Else
        'Do nothing
    End Select
Exit Sub
'Error Handler
tlbMaster_ButtonClick_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Public Sub SessionSaveMenuRoutine()
'Created by Robin G Brown, 1/9/97
'Save all the forms and session data
'sub declares
Const conSub = "SessionSaveMenuRoutine"
Dim strFileAndPath As String
Dim strOriginalName As String
Dim frmCurrent As Form
Dim lngFormNumber As Long
Dim strWindow As String
    'Error Trap
    On Error GoTo SessionSaveMenuRoutine_ErrorHandler
    With frmMaster.dlgMaster
        .DialogTitle = "Save Session"
        .CancelError = True
        .Flags = cdlOFNHideReadOnly Or cdlOFNCreatePrompt
        .Filter = conAnyFilter & "|" & conTriSessionFilter
        .FilterIndex = 2
        .filename = "Session"
        .ShowSave
        strFileAndPath = .filename
    End With
    'Save the data and formatting to a section file - sffSession
    Set sffSession = CreateObject("Sectionfile.SecFile")
    With sffSession
        'Details of all open files
        lngFormNumber = 0
        For Each frmCurrent In Forms
            If TypeOf frmCurrent Is frmGridReport Then
                If frmCurrent.IsDirty = True Then
                    frmCurrent.FileSaveMenuRoutine
                End If
                lngFormNumber = lngFormNumber + 1
                strWindow = "WINDOW" & lngFormNumber
                .AddSection strWindow
                .SetSectionSetting strWindow, "file", frmCurrent.SaveFileName
                .SetSectionSetting strWindow, "type", "GRID"
                'Window size and position
                .SetSectionSetting strWindow, "state", frmCurrent.WindowState
                .SetSectionSetting strWindow, "left", frmCurrent.Left
                .SetSectionSetting strWindow, "top", frmCurrent.Top
                .SetSectionSetting strWindow, "height", frmCurrent.Height
                .SetSectionSetting strWindow, "width", frmCurrent.Width
            End If
        Next
        If FormInstances("frmChartReport") > 0 Then
            If frmChartReport.IsDirty = True Then
                frmChartReport.FileSaveMenuRoutine
            End If
            lngFormNumber = lngFormNumber + 1
            strWindow = "WINDOW" & lngFormNumber
            .AddSection strWindow
            .SetSectionSetting strWindow, "file", frmChartReport.SaveFileName
            .SetSectionSetting strWindow, "type", "CHART"
            'Window size and position
            .SetSectionSetting strWindow, "state", frmChartReport.WindowState
            .SetSectionSetting strWindow, "left", frmChartReport.Left
            .SetSectionSetting strWindow, "top", frmChartReport.Top
            .SetSectionSetting strWindow, "height", frmChartReport.Height
            .SetSectionSetting strWindow, "width", frmChartReport.Width
        End If
        .AddSection "FILES"
        .SetSectionSetting "FILES", "number", lngFormNumber
        .FileSave strFileAndPath
    End With
    'Clear references
    Set frmCurrent = Nothing
Exit Sub
SessionSaveMenuRoutine_ErrorHandler:
    Exit Sub
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbOKOnly, conModule & ":" & conSub
    End Select
    Exit Sub
End Sub

Public Sub SessionFileOpenMenuRoutine(ByVal strOpenFileName As String)
'Created by Robin G Brown, 19/5/97
'Retrieve the formatting and data from a file
'Sub Declares
Const conSub = "FileOpen"
Dim lngNumberOfWindows As Long
Dim lngWindowCounter As Long
Dim strFileName As String
Dim strFileType As String
Dim lngPositionItem As Long
Dim strWindow As String
Dim frmNew As frmGridReport
    'Error Trap
    On Error GoTo SessionFileOpenMenuRoutine_ErrorHandler
    Set sffSession = CreateObject("Sectionfile.SecFile")
    With sffSession
        'Open the file
        .FileOpen strOpenFileName
        lngNumberOfWindows = 0
        lngNumberOfWindows = .GetSectionSetting("FILES", "number")
        For lngWindowCounter = 1 To lngNumberOfWindows
            strWindow = "WINDOW" & lngWindowCounter
            strFileName = .GetSectionSetting(strWindow, "file")
            strFileType = .GetSectionSetting(strWindow, "type")
            Select Case strFileType
            Case "CHART"
                Err.Clear
                Load frmChartReport
                If Err.Number <> 0 Then
                    Exit Sub
                End If
                Call frmChartReport.GraphFileOpenMenuRoutine(strFileName)
                'Window size and position
                frmChartReport.WindowState = .GetSectionSetting(strWindow, "state")
                frmChartReport.Left = .GetSectionSetting(strWindow, "left")
                frmChartReport.Top = .GetSectionSetting(strWindow, "top")
                frmChartReport.Height = .GetSectionSetting(strWindow, "height")
                frmChartReport.Width = .GetSectionSetting(strWindow, "width")
                frmChartReport.Show
            Case "GRID"
                Set frmNew = New frmGridReport
                Load frmNew
                Call frmNew.GridFileOpenMenuRoutine(strFileName)
                'Window size and position
                frmNew.WindowState = .GetSectionSetting(strWindow, "state")
                frmNew.Left = .GetSectionSetting(strWindow, "left")
                frmNew.Top = .GetSectionSetting(strWindow, "top")
                frmNew.Height = .GetSectionSetting(strWindow, "height")
                frmNew.Width = .GetSectionSetting(strWindow, "width")
                frmNew.Show
                Set frmNew = Nothing
            Case Else
                'Ignore
            End Select
        Next
    End With
Exit Sub
'Error Handler
SessionFileOpenMenuRoutine_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbOKOnly, conModule & ":" & conSub
    End Select
    Exit Sub
End Sub

Property Get NextFormNumber() As Long
'Created by Robin G Brown, 21/8/97
'Return the next valid number for a form
'Property Declares
Const conProperty = "NextFormNumber"
Static lngFormNumber As Long
'Error Trap
    On Error Resume Next
    lngFormNumber = lngFormNumber + 1
    NextFormNumber = lngFormNumber
End Property


