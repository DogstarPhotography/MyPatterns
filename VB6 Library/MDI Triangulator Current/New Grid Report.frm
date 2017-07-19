VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Object = "{CB5103A1-C04E-11CF-91F7-C2863C385E30}#2.0#0"; "Vsflex2.ocx"
Begin VB.Form frmGridReport 
   Caption         =   "Report"
   ClientHeight    =   5355
   ClientLeft      =   1155
   ClientTop       =   1560
   ClientWidth     =   8340
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
   Icon            =   "New Grid Report.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5355
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin vsFlexLib.vsFlexArray grdReport 
      Height          =   2415
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5655
      _Version        =   131072
      _ExtentX        =   9975
      _ExtentY        =   4260
      _StockProps     =   228
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
      Appearance      =   1
      FixedRows       =   2
      FixedCols       =   2
      FocusRect       =   0
      HighLight       =   0
      GridColor       =   8421504
      MergeCells      =   4
      PictureType     =   1
      AllowUserResizing=   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   2580
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin ComctlLib.ImageList imlIcons 
      Left            =   60
      Top             =   2580
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "New Grid Report.frx":0442
            Key             =   "singlecol"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "New Grid Report.frx":075C
            Key             =   "multicol"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFileRoot 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&New Report"
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "New &Graph"
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Open..."
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Save..."
         Index           =   4
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Save Session"
         Index           =   5
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Print"
         Index           =   6
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Close"
         Index           =   7
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   8
      End
   End
   Begin VB.Menu mnuWindowRoot 
      Caption         =   "&Window"
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
         Caption         =   "&Windows"
         WindowList      =   -1  'True
      End
   End
   Begin VB.Menu mnuReportRoot 
      Caption         =   "&Report"
      Begin VB.Menu mnuReport 
         Caption         =   "&Expand"
         Index           =   1
      End
      Begin VB.Menu mnuReport 
         Caption         =   "&Collapse"
         Index           =   2
      End
      Begin VB.Menu mnuReport 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuReport 
         Caption         =   "&Graph Column"
         Index           =   4
      End
      Begin VB.Menu mnuReport 
         Caption         =   "&Account Display..."
         Index           =   5
      End
      Begin VB.Menu mnuReport 
         Caption         =   "&ShowSource"
         Index           =   6
      End
   End
   Begin VB.Menu mnuHelpRoot 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmGridReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Robin G Brown, 2nd May 1997
'-----------------------------------------------------------------------------
'   This form contains code for _
    showing a report grid that displays data held in a TriData object
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "frmGridReport"
Private intCounter As Integer
Private lngReturn As Long
Private intColCounter As Integer
Private intRowCounter As Integer
Public dtaDataArray As TriData
Private booshowChart As Boolean
Private booClickOnGrid As Boolean
'Hold a reference for saving the report window
Private sffReport As SecFile
'Variables to store saved status of this form
Private booIsDirty As Boolean
Private strSaveFileName As String

Private Sub Form_Initialize()
'Created by Robin G Brown, 27/5/97
'Initialize the data array
'Sub Declares
Const conSub = "Form_Initialize"
Dim lngInstance As Long
    'Error Trap
    On Error Resume Next
    'Make sure the datagrid array is set up
    Set dtaDataArray = New TriData
    If Err.Number <> 0 Then
        MsgBox "Data failed to initialise due to error : " & Err.Description, vbExclamation, conModule & " : " & conSub
        Exit Sub
    End If
    'Set up dtaDataArray
    Set dtaDataArray.FlexArray = grdReport
    dtaDataArray.Name = "Report" & frmMaster.NextFormNumber
    'Create a section file object
    Set sffReport = CreateObject("SectionFile.SecFile")
    'Set any other variables
    booClickOnGrid = False
    booIsDirty = True
End Sub

Private Sub Form_Load()
'Created by Robin G Brown, 2/5/97
'Default Behaviour
'Sub Declares
Const conSub = "Form_Load"
    'Error Trap
    On Error GoTo Form_Load_ErrorHandler
    'Set the caption
    Me.Caption = dtaDataArray.Name
    Me.WindowState = vbNormal
    'Data will be set by whatever causes this form to load
Exit Sub
'Error Handler
Form_Load_ErrorHandler:
    Exit Sub
End Sub

Private Sub Form_Resize()
'Created by Robin G Brown, 2/5/97
'Default Behaviour
'Sub Declares
Const conSub = "Form_Resize"
    'Error Trap
    On Error GoTo Form_Resize_ErrorHandler
    'Resize the grid to fill the form
    If Me.WindowState <> 1 Then
        grdReport.Height = Me.Height - 420
        grdReport.Width = Me.Width - 120
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

Public Sub DataRandomInitialize(Optional ByVal varStartYear As Variant)
'Created by Robin G Brown, 27/5/97
'Initialise the data using the given start year and creating random data
'Sub Declares
Const conSub = "DataRandomInitialize"
Dim intStartYear As Integer
    'Error Trap
    On Error Resume Next
    If IsMissing(varStartYear) Then
        'Create a default start year
        intStartYear = 1990
    Else
        intStartYear = CInt(varStartYear)
        If Err.Number <> 0 Then
            intStartYear = 1990
            MsgBox "Error : " & Err.Description, vbExclamation, conModule & " : " & conSub
            Err.Clear
        End If
    End If
    'Create some random data
    Call dtaDataArray.CreateData(intStartYear)
    'Fill the grid from dtaDataArray
    Call dtaDataArray.FillFlexArrayWithYearlyData
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Created by Robin G Brown, 1/9/97
'Check to see if this form requires saving
'Function Declares
Const conFunction = "Form_Unload"
Dim lngTempValue As Long
    'Errors 'disabled' to ensure that form unloads
    On Error Resume Next
    If booIsDirty = True Then
        lngReturn = MsgBox("Grid not saved. Save before closing?", vbYesNoCancel, App.ProductName)
        Select Case lngReturn
        Case vbYes
            'Save
            Call FileSaveMenuRoutine
        Case vbCancel
            'Cancel close
            Cancel = 1
        Case Else
            'Carry on without saving
        End Select
    End If
    'CODE_HERE
End Sub

Private Sub grdReport_Click()
'Created by Robin G Brown, 28/8/97
'Expand /collapse, depending on where the click took place
'Sub Declares
Const conSub = "grdReport_DblClick"
    'Error Trap
    On Error GoTo grdReport_DblClick_ErrorHandler
    If grdReport.MouseCol > grdReport.Cols Or grdReport.MouseRow > grdReport.Rows Then
        'Ignore if the click is outside the 'grid' area
        Exit Sub
    End If
    If grdReport.MouseCol < grdReport.FixedCols Then
        'Expand/collapse display
        Select Case grdReport.MouseCol
        Case 0
            If grdReport.TextMatrix(2, 1) = "" Then
                Call dtaDataArray.FillFlexArrayWithQuarterlyData
            Else
                Call dtaDataArray.FillFlexArrayWithYearlyData
            End If
        Case 1
            If grdReport.TextMatrix(2, 1) = "" Then
                Call dtaDataArray.FillFlexArrayWithQuarterlyData
            ElseIf grdReport.TextMatrix(2, 2) = "" Then
                Call dtaDataArray.FillFlexArrayWithMonthlyData
            Else
                Call dtaDataArray.FillFlexArrayWithQuarterlyData
            End If
        Case Else
            If grdReport.TextMatrix(2, 2) = "" Then
                Call dtaDataArray.FillFlexArrayWithMonthlyData
            Else
                Call dtaDataArray.FillFlexArrayWithQuarterlyData
            End If
        End Select
    End If
Exit Sub
'Error Handler
grdReport_DblClick_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Private Sub grdReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Created by Robin G Brown, 13/5/97
'Make sure that gfgcCurrent is set
'Sub Declares
    'Error Trap
    On Error Resume Next
    booshowChart = False
    'Set the mouse col and row for the grid
    dtaDataArray.GridMouseCol = grdReport.MouseCol
    If Err <> 0 Then
        'Invalid mouse position!
        Err.Clear
    End If
    dtaDataArray.GridMouserow = grdReport.MouseRow
    If Err <> 0 Then
        'Invalid mouse position!
        Err.Clear
    End If
    'Popup the grid menu of frmmaster if the right mouse button is being used
    If Button = vbRightButton Then
        booClickOnGrid = True
        PopupMenu Menu:=mnuReportRoot, defaultmenu:=mnuReport(4)
    End If
End Sub

Private Sub grdReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Created by Robin G Brown, 23/5/97
'Start a drag/drop operation
'Sub Declares
    'Error Trap
    On Error Resume Next
    'Check that the mouse is in the right position - RGB/17/9/97
    If Button = vbLeftButton _
    And grdReport.MouseCol >= grdReport.FixedCols _
    And grdReport.MouseRow <= grdReport.Rows _
    And grdReport.MouseCol <= grdReport.Cols _
    And grdReport.MouseRow > 0 _
    And grdReport.MouseCol > 0 _
    Then
        'Set the drag icon
        If frmMaster.stbStatus.Panels("panYears").Tag = "OFF" Then
            'Single col drag
            grdReport.DragIcon = imlIcons.ListImages("singlecol").Picture
        Else
            'Multiple column drag
            grdReport.DragIcon = imlIcons.ListImages("multicol").Picture
        End If
        grdReport.Drag vbBeginDrag
        gbooColumnDragInProgress = True
    End If
End Sub

Private Sub grdReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Created by Robin G Brown, 13/5/97
'Sub Declares
    'Error Trap
    On Error Resume Next
    'Cancel a drag/drop operation
    grdReport.Drag vbCancel
    gbooColumnDragInProgress = False
    booshowChart = False
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
    Case 4
        Call FileSaveMenuRoutine
    Case 5
        Call frmMaster.SessionSaveMenuRoutine
    Case 6
        Call FilePrintMenuRoutine
    Case 7
        Unload frmMaster.ActiveForm
    Case 8
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

Private Sub mnuReportRoot_Click()
'Created by Robin G Brown, 29/8/97
'Disable any menus that are not appropriate
    'Error Trap
    On Error Resume Next
    mnuReport(4).Enabled = False
    If booClickOnGrid = True And grdReport.MouseCol > grdReport.FixedCols Then
        mnuReport(4).Enabled = True
        booClickOnGrid = False
    End If
End Sub

Private Sub mnuReport_Click(Index As Integer)
'Created by Robin G Brown, 29/8/97
'Call appropriate
'Sub Declares
Const conSub = "mnuReport_Click"
    'Error Trap
    On Error GoTo mnuReport_Click_ErrorHandler
    'Call the selected routine passing a reference to the current form
    Select Case Index
    Case 1
        'Expand the report to the next level
        If grdReport.TextMatrix(2, 1) = "" Then
            Call dtaDataArray.FillFlexArrayWithQuarterlyData
        Else
            Call dtaDataArray.FillFlexArrayWithMonthlyData
        End If
    Case 2
        'Collapse the report to the previous level
        If grdReport.TextMatrix(2, 2) <> "" Then
            Call dtaDataArray.FillFlexArrayWithQuarterlyData
        Else
            Call dtaDataArray.FillFlexArrayWithYearlyData
        End If
    Case 4
        'Add this column to the graph
        If FormInstances("frmChartReport") = 0 Then
            Beep
            Exit Sub
        End If
        'Graph the appropriate column(s)
        Call frmChartReport.GraphColumns(Me.dtaDataArray)
    Case 5
        Call dtaDataArray.ShowDisplaySettings
    Case 6
        MsgBox "Source data not available.", vbOKOnly, App.ProductName
    Case Else
        'Default - do nothing
    End Select
Exit Sub
mnuReport_Click_ErrorHandler:
    MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    Exit Sub
End Sub

Private Sub mnuWindow_Click(Index As Integer)
'Created by Robin G Brown, 28/8/97
'Call appropriate StandardMDIArrange in modCommonfrmmasternus module
'Sub Declares
Const conSub = "mnuWindow_Click"
    'Error Trap
    On Error GoTo mnuWindow_Click_ErrorHandler
    'Call the selected routine passing a reference to the current form
    Select Case Index
    Case 1
        frmMaster.Arrange vbTileHorizontal
    Case 2
        frmMaster.Arrange vbTileVertical
    Case 3
        frmMaster.Arrange vbCascade
    Case 4
        frmMaster.Arrange vbArrangeIcons
    Case Else
        'Default - do nothing
    End Select
Exit Sub
mnuWindow_Click_ErrorHandler:
    MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    Exit Sub
End Sub

Public Sub GridFileOpenMenuRoutine(ByVal strOpenFileName As String)
'Created by Robin G Brown, 28/8/97
'Retrieve the contents of a file into the current grid
'sub declares
    'Error Trap
    On Error GoTo GridFileOpenMenuRoutine_ErrorHandler
    strSaveFileName = strOpenFileName
    Set dtaDataArray.FlexArray = grdReport
    Call dtaDataArray.FileOpen(sffReport, strSaveFileName)
    Call dtaDataArray.FillFlexArrayWithYearlyData
    Me.Caption = StripFileExtension(GetFileNameFromPath(strSaveFileName))
    dtaDataArray.Name = Me.Caption
    booIsDirty = False
Exit Sub
GridFileOpenMenuRoutine_ErrorHandler:
    Exit Sub
End Sub

Public Sub FileSaveMenuRoutine()
'Created by Robin G Brown, 19/5/97
'Save this form and all associated data to a section file
'sub declares
Const conSub = "FileSaveMenuRoutine"
Dim strOriginalName As String
    'Error Trap
    On Error GoTo FileSaveMenuRoutine_ErrorHandler
    With frmMaster.dlgMaster
        .DialogTitle = "Save Grid"
        .CancelError = True
        .Flags = cdlOFNHideReadOnly Or cdlOFNCreatePrompt
        .Filter = conAnyFilter & "|" & conTriReportFilter
        .FilterIndex = 2
        .filename = Me.Caption
        .ShowSave
        strSaveFileName = .filename
    End With
    'Save the data and formatting to a section file - sffReport
    'Pass sffReport over to dtaDataArray instructing it to set up and save sffReport
    Call dtaDataArray.FileSave(sffReport, strSaveFileName)
    'Then set the caption of the form and the name of the data array
    strOriginalName = Me.Caption
    Me.Caption = StripFileExtension(GetFileNameFromPath(strSaveFileName))
    Me.dtaDataArray.Name = Me.Caption
    'Then set the report name in the chart
    If FormInstances("frmChartReport") > 0 Then
        Call frmChartReport.ChangeReportNameInGraph(strOriginalName, Me.Caption)
    End If
    booIsDirty = False
Exit Sub
FileSaveMenuRoutine_ErrorHandler:
    Exit Sub
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbOKOnly, conModule & ":" & conSub
    End Select
    Exit Sub
End Sub

Public Sub FilePrintMenuRoutine()
'Created by Robin G Brown, 17/9/97
'Print the report to the default printer
'sub declares
Const conSub = "FilePrintMenuRoutine"
Dim lngPicWidth As Long
Dim lngPicHeight As Long
Dim sngScaleFactorPortrait As Single
Dim sngScaleFactorLandscape As Single
Dim sngScaleFactorActual As Single
    'Error Trap
    On Error GoTo FilePrintMenuRoutine_ErrorHandler
    'Printer dialog
    With frmMaster.dlgMaster
        'PrinterDefault needs to be set true _
            so that we don't have to do a lot of difficult GDI work! - RGB/17/9/97
        .PrinterDefault = True
        .ShowPrinter
        Printer.Copies = .Copies
    End With
    'Print
    lngPicWidth = grdReport.Picture.Width
    lngPicHeight = grdReport.Picture.Height
    'Work out a scaling factor for displaying the grid
    sngScaleFactorActual = 1#
    With Printer
        'Try it portrait first
        .Orientation = vbPRORPortrait
        sngScaleFactorPortrait = 1#
        'It doesn't matter which is checked first, _
            height or width, as long as sngScaleFactorPortrait is  taken into account
        If (lngPicHeight * sngScaleFactorPortrait) > (.Height - 800) Then
            sngScaleFactorPortrait = sngScaleFactorPortrait * ((.Height - 800) / lngPicHeight)
        Else
            sngScaleFactorPortrait = sngScaleFactorPortrait * 1#
        End If
        If (lngPicWidth * sngScaleFactorPortrait) > (.Width - 600) Then
            sngScaleFactorPortrait = sngScaleFactorPortrait * ((.Width - 600) / lngPicWidth)
        Else
            sngScaleFactorPortrait = sngScaleFactorPortrait * 1#
        End If
        'Then landscape
        .Orientation = vbPRORLandscape
        sngScaleFactorLandscape = 1#
        'It doesn't matter which is checked first, _
            height or width, as long as sngScaleFactorLandscape is  taken into account
        If (lngPicHeight * sngScaleFactorLandscape) > (.Height - 800) Then
            sngScaleFactorLandscape = sngScaleFactorLandscape * ((.Height - 800) / lngPicHeight)
        Else
            sngScaleFactorLandscape = sngScaleFactorLandscape * 1#
        End If
        If (lngPicWidth * sngScaleFactorLandscape) > (.Width - 600) Then
            sngScaleFactorLandscape = sngScaleFactorLandscape * ((.Width - 600) / lngPicWidth)
        Else
            sngScaleFactorLandscape = sngScaleFactorLandscape * 1#
        End If
        'now choose the scaling factor that has the least reduction
        If sngScaleFactorLandscape >= sngScaleFactorPortrait Then
            .Orientation = vbPRORLandscape
            sngScaleFactorActual = sngScaleFactorLandscape
        Else
            .Orientation = vbPRORPortrait
            sngScaleFactorActual = sngScaleFactorPortrait
        End If
    End With
    'Now draw the picture on the printer
    Printer.PaintPicture grdReport.Picture, 0, 0, _
        grdReport.Picture.Width * sngScaleFactorActual, _
        grdReport.Picture.Height * sngScaleFactorActual
    'Print
    Printer.EndDoc
Exit Sub
FilePrintMenuRoutine_ErrorHandler:
    Exit Sub
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbOKOnly, conModule & ":" & conSub
    End Select
    Exit Sub
End Sub

Private Sub Form_Terminate()
'Created by Robin G Brown, 12/8/97
'Clear any references
    On Error Resume Next
    Set dtaDataArray = Nothing
    Set sffReport = Nothing
End Sub

Property Get SaveFileName() As String
'Created by Robin G Brown, 1/9/97
'Property Declares
Const conProperty = "SaveFileName"
'Error Trap
    On Error Resume Next
    SaveFileName = strSaveFileName
End Property

Property Get IsDirty() As Boolean
'Created by Robin G Brown, 1/9/97
'Property Declares
Const conProperty = "IsDirty"
'Error Trap
    On Error Resume Next
    IsDirty = booIsDirty
End Property


