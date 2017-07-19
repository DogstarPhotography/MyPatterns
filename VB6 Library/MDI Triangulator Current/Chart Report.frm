VERSION 5.00
Object = "{02B5E320-7292-11CF-93D5-0020AF99504A}#1.0#0"; "MSCHART.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "comdlg32.ocx"
Begin VB.Form frmChartReport 
   BackColor       =   &H8000000A&
   Caption         =   "Chart"
   ClientHeight    =   1305
   ClientLeft      =   1155
   ClientTop       =   1275
   ClientWidth     =   2520
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
   Icon            =   "Chart Report.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1305
   ScaleWidth      =   2520
   WindowState     =   2  'Maximized
   Begin MSChartLib.MSChart graReport 
      Height          =   975
      Left            =   0
      OleObjectBlob   =   "Chart Report.frx":0442
      TabIndex        =   0
      Top             =   0
      Width           =   1515
   End
   Begin MSComDlg.CommonDialog dlgChart 
      Left            =   420
      Top             =   180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin VB.Menu mnuFileRoot 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&New Report"
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "New &Graph"
         Enabled         =   0   'False
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
   Begin VB.Menu mnuGraph 
      Caption         =   "&Graph"
      Begin VB.Menu mnuGraphDrillDown 
         Caption         =   "&Show Source"
      End
      Begin VB.Menu mnuGraphClear 
         Caption         =   "Clear"
      End
      Begin VB.Menu mnuGraphColour 
         Caption         =   "&Colour"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelpRoot 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmChartReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Robin G Brown, 2nd May 1997
'-----------------------------------------------------------------------------
'   This form contains code for _
    showing a report grid
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "frmChartReport"
Private Const conSeperator = ","
Private intCounter As Integer
Private lngReturn As Long
Private booChartNew As Boolean
Private intCurrentSeries As Integer
Private booMonoColour As Boolean
'Hold a reference for saving the graph window
Private sffGraph As SecFile
'Variables to store saved status of this form
Private booIsDirty As Boolean
Private strSaveFileName As String

Public Sub ChangeReportNameInGraph(ByVal strReportName As String, ByVal strReplace As String)
'Created by Robin G Brown, 1/9/97
'Change all occurences of strReportName in the Graph to strReplace
'Sub Declares
Const conSub = "ChangeReportNameInGraph"
Dim lngColCounter As Long
Dim lngNameStartPos As Long
Dim strColumnLabel As String
    'Error Trap
    On Error GoTo ChangeReportNameInGraph_ErrorHandler
    With graReport.DataGrid
        For lngColCounter = 1 To .ColumnCount
            strColumnLabel = .ColumnLabel(lngColCounter, 1)
            If InStr(strColumnLabel, strReportName) > 0 Then
                strColumnLabel = ReplaceStringInString(strColumnLabel, strReportName, strReplace)
                .ColumnLabel(lngColCounter, 1) = strColumnLabel
            End If
        Next
    End With
    booIsDirty = True
Exit Sub
'Error Handler
ChangeReportNameInGraph_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Public Sub ColouriseChart()
'Created by Robin G Brown, 20/8/97
'Make all the chart series the right colour and a standard width
'Sub Declares
Dim intChartColCounter As Integer
Dim intCurrentAccount As Integer
Dim intRed As Integer
Dim intGreen As Integer
Dim intBlue As Integer
    'Error Trap
    On Error Resume Next
    With graReport
        If .chartType = VtChChartType2dLine Then
            For intChartColCounter = 1 To .Plot.SeriesCollection.Count
                    'Set the colour and the width
                    Call SetNextPen(.Plot.SeriesCollection(intChartColCounter).Pen, intChartColCounter, intChartColCounter Mod 8)
            Next
        ElseIf .chartType = VtChChartType3dLine Then
            'Use a solid pen
            For intChartColCounter = 1 To .Plot.SeriesCollection.Count
                'Always use the same pen for 3D charts
                Call SetNextPen(.Plot.SeriesCollection(intChartColCounter).Pen, 1, intChartColCounter Mod 8)
            Next
        Else
            'All other charts use a brush
            For intChartColCounter = 1 To .Plot.SeriesCollection.Count
                'Set the colour and the width
                Call SetNextBrush(.Plot.SeriesCollection(intChartColCounter).DataPoints(-1).Brush, intChartColCounter Mod 8)
'                With .Plot.SeriesCollection(intChartColCounter)
'                    .DataPoints(-1).Brush.FillColor.Set 0, 0, 0
'                End With
            Next
        End If
    End With
End Sub

Public Sub ClearChart()
'Created by Robin G Brown, 21/8/97
'Clear the chart of any information
'Sub Declares
    'Error Trap
    On Error Resume Next
    Call RemoveAllColumns(graReport)
    'Add something back so that there are axes showing
    Call DrawBlankAxes(graReport)
    booChartNew = True
End Sub

Private Sub Form_Initialize()
On Error Resume Next
    booMonoColour = True
    booChartNew = True
    booIsDirty = False
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
        lngReturn = MsgBox("Graph not saved. Save before closing?", vbYesNoCancel, App.ProductName)
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
End Sub

Private Sub graReport_DragDrop(Source As Control, X As Single, Y As Single)
'Created by Robin G Brown, 28/5/97
'Handle columns being dropped onto the chart
'Sub Declares
Const conSub = "graReport_DragDrop"
Dim dtaSourceArray As TriData
    'Error Trap
    On Error GoTo graReport_DragDrop_ErrorHandler
    'Check that the source is a flexarray, do not handle any other sources
    If Not TypeOf Source Is vsFlexArray Then
        Beep
        Exit Sub
    End If
    'Cal the appropriate routine, passing over the DataArray
    Call GraphColumns(Source.Parent.dtaDataArray)
    booIsDirty = True
Exit Sub
'Error Handler
graReport_DragDrop_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Set dtaSourceArray = Nothing
    Exit Sub
End Sub

Public Sub GraphColumns(ByRef dtaSourceArray As TriData)
'Created by Robin G Brown, 28/5/97
'Add column(s) to the chart by calling methods of dtaSourceArray
'Sub Declares
Const conSub = "GraphColumns"
Dim strAccountName As String
    'Error Trap
    On Error GoTo GraphColumns_ErrorHandler
    'Check for comparing different accounts on the same graph
    If booChartNew = False Then
        graReport.Column = 1
        With dtaSourceArray
            strAccountName = .AccountName(.GetAccountFromColumn(.GridMouseCol))
        End With
        If InStr(graReport.ColumnLabel, strAccountName) = 0 Then
            If MsgBox("You are attempting to compare different accounts, continue?", vbOKCancel, App.ProductName) = vbCancel Then
                Exit Sub
            End If
        End If
    End If
    'Call the appropriate functions
    Set dtaSourceArray.Chart = Me.graReport
    'check that the number of rows on the grid matches the number of rows in the graph
    If dtaSourceArray.DataRows <> graReport.RowCount And booChartNew = False Then
        lngReturn = MsgBox("New data cannot be compared with currently displayed data, clear current data?", vbOKCancel, App.ProductName)
        If lngReturn = vbCancel Then
            Set dtaSourceArray = Nothing
            Exit Sub
        Else
            'This will set booChartNew = True and clear any current bits on the chart
            Call ClearChart
        End If
    End If
    If frmMaster.stbStatus.Panels("panYears").Tag = "OFF" Then
        'Has the chart been used before?
        If booChartNew = True Then
            Call dtaSourceArray.GraphOnColumn(dtaSourceArray.GridMouseCol)
            booChartNew = False
            Call ShowLegend(Me.graReport)
        Else
            Call dtaSourceArray.AddColumnToGraph(dtaSourceArray.GridMouseCol)
        End If
        Call Me.ColouriseChart
    Else
        'Multiple account - check to see if the user wants to add or replace multiple columns
        'Has the chart been used before?
        If booChartNew = True Then
            Call dtaSourceArray.GraphOnAccount(dtaSourceArray.GridMouseCol)
            booChartNew = False
            Call ShowLegend(Me.graReport)
        Else
            If MsgBox("Replace current columns? Click Yes to replace, No to add.", vbYesNo, App.ProductName) = vbNo Then
                Call dtaSourceArray.AddMultipleColumnsToGraph(dtaSourceArray.GridMouseCol)
            Else
                Call dtaSourceArray.GraphOnAccount(dtaSourceArray.GridMouseCol)
            End If
        End If
        Call Me.ColouriseChart
    End If
    Call PositionLegendAndPlot(Me.graReport, conAlignLegendBottomCentre, 90, 10)
    'Call MaximizePlot(Me.graReport)
    'Clear the reference
    Set dtaSourceArray = Nothing
    graReport.Refresh
Exit Sub
'Error Handler
GraphColumns_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Set dtaSourceArray = Nothing
    Exit Sub
End Sub

Private Sub Form_Activate()
'Created by Robin G Brown, 27/5/97
'Activate the chart type selection drop down
'Sub Declares
Const conSub = "Form_Activate"
    'Error Trap
    On Error Resume Next
    Select Case graReport.chartType
    Case VtChChartType2dLine
        '2D Line
        frmMaster.cmbChartType.ListIndex = 0
    Case VtChChartType3dLine
        '3D Line
        frmMaster.cmbChartType.ListIndex = 1
    Case VtChChartType2dArea
        '2D Area
        frmMaster.cmbChartType.ListIndex = 2
    Case VtChChartType3dArea
        '3D Area
        frmMaster.cmbChartType.ListIndex = 3
    Case VtChChartType2dBar
        '2D Bar
        frmMaster.cmbChartType.ListIndex = 4
    Case VtChChartType3dBar
        '3D Bar
        frmMaster.cmbChartType.ListIndex = 5
    Case VtChChartType2dStep
        '2D Step
        frmMaster.cmbChartType.ListIndex = 6
    Case VtChChartType3dStep
        '3D Step
        frmMaster.cmbChartType.ListIndex = 7
    Case Else
        frmMaster.cmbChartType.ListIndex = 0
    End Select
    frmMaster.cmbChartType.Enabled = True
End Sub

Private Sub Form_Deactivate()
'Created by Robin G Brown, 27/5/97
'Activate the chart type selection drop down
'Sub Declares
Const conSub = "Form_Deactivate"
    'Error Trap
    On Error Resume Next
    frmMaster.cmbChartType.ListIndex = 0
    frmMaster.cmbChartType.Enabled = False
End Sub

Private Sub Form_Load()
'Created by Robin G Brown, 20/8/97
'Sub Declares
Const conSub = "Form_Load"
    'Error Trap
    On Error GoTo Form_Load_ErrorHandler
    'When this form opens it should be 'normal' sized - which is quite small for this form
    Me.Caption = "Chart" & frmMaster.NextFormNumber
    Me.WindowState = vbNormal
    Me.Height = 1800
    Me.Width = 2000
    Me.Show
    Call DrawBlankAxes(graReport)
    booChartNew = True
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
'Created by Robin G Brown, 2/5/97
'Default Behaviour
'Sub Declares
Const conSub = "Form_Resize"
    'Error Trap
    On Error GoTo Form_Resize_ErrorHandler
    If Me.WindowState <> 1 Then
        'Resize the grid to fill the form
        graReport.Height = Me.Height - 400
        graReport.Width = Me.Width - 120
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

Private Sub graReport_KeyUp(KeyCode As Integer, Shift As Integer)
'Created by Robin G Brown, 20/8/97
'If DEL has been pressed, delete the current series
'Sub Declares
Const conSub = "graReport_KeyUp"
    'Error Trap
    On Error Resume Next
    If KeyCode = vbKeyDelete And intCurrentSeries > 0 Then
        'Use RemoveColumn to remove the column
        Call RemoveColumn(graReport, intCurrentSeries)
        Call ColouriseChart
    End If
    If graReport.DataGrid.ColumnCount = 0 Then
        booChartNew = True
        Call DrawBlankAxes(graReport)
    End If
End Sub

'Private Sub graReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
''Created by Robin G Brown, 13/5/97
''popup menu
''Sub Declares
'    'Error Trap
'    On Error Resume Next
'    'Work out where the mouse was pressed?
'    'Popup the graph menu of frmmaster if the right mouse button is being used
'    If Button = vbRightButton Then
'        PopupMenu Menu:=mnuGraph, defaultmenu:=mnuGraphDrillDown
'    End If
'End Sub

Private Sub graReport_SeriesActivated(Series As Integer, MouseFlags As Integer, Cancel As Integer)
'Created by Robin G Brown, 4/9/97
'popup menu on a double click
'Sub Declares
    'Error Trap
    On Error Resume Next
    'Work out where the mouse was pressed?
    intCurrentSeries = Series
    'Popup the graph menu of this form
    PopupMenu Menu:=mnuGraph, defaultmenu:=mnuGraphDrillDown
End Sub

Private Sub graReport_SeriesSelected(Series As Integer, MouseFlags As Integer, Cancel As Integer)
On Error Resume Next
    intCurrentSeries = Series
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

Private Sub mnuGraphClear_Click()
'Created by Robin G Brown, 21/8/97
'Clear the chart
    'Error Trap
    On Error Resume Next
    Call ClearChart
End Sub

Private Sub mnuGraphColour_Click()
'Created by Robin G Brown, 21/8/97
'Switch colour for the chart on or off
    'Error Trap
    On Error Resume Next
    'Set colour on/off
    If mnuGraphColour.Checked = True Then
        booMonoColour = False
        mnuGraphColour.Checked = False
    Else
        booMonoColour = True
        mnuGraphColour.Checked = True
    End If
    'Apply to the chart
    Call ColouriseChart
End Sub

Private Sub mnuGraphDrillDown_Click()
'Created by Robin G Brown, 26/8/97
'Show the grid, highlighting the appropriate column
'Sub Declares
Dim strLabel As String
Dim frmCurrent As Form
    'Error Trap
    On Error Resume Next
    'No line is selected until after a click event on the graph _
        so there is likely to be either no line selected or the wrong line selected
    'Try anyway
    If intCurrentSeries > 0 Then
        'Get the label for the selected series
        strLabel = graReport.DataGrid.ColumnLabel(intCurrentSeries, 1)
        For Each frmCurrent In Forms
            'Match the current windows to the selected series
            If InStr(strLabel, frmCurrent.Caption) > 0 Then
                'Push the selected window to the top
                frmCurrent.ZOrder 0
                Exit Sub
            End If
        Next
        'Not found
        MsgBox "Source report not found.", vbOKOnly, App.ProductName
    End If
End Sub

Public Sub FileSaveMenuRoutine()
'Created by Robin G Brown, 19/5/97
'Save this form and all associated data to a section file
'sub declares
Const conSub = "FileSaveMenuRoutine"
Dim strDataLine As String
Dim lngColCounter As Long
Dim lngRowCounter As Long
Dim lngColCount As Long
Dim lngRowcount As Long
Dim dblDataPoint As Double
Dim intNullFlag As Integer
    'Error Trap
    On Error GoTo FileSaveMenuRoutine_ErrorHandler
    With frmMaster.dlgMaster
        .DialogTitle = "Save Graph"
        .CancelError = True
        .Flags = cdlOFNHideReadOnly Or cdlOFNCreatePrompt
        .Filter = conAnyFilter & "|" & conTriGraphFilter
        .FilterIndex = 2
        .filename = Me.Caption
        .ShowSave
        strSaveFileName = .filename
    End With
    'Save the data and formatting to a section file - sffGraph
    Set sffGraph = New SecFile
    With sffGraph
        'Window size and position
        .AddSection "WINDOW"
        .SetSectionSetting "WINDOW", "state", Me.WindowState
        .SetSectionSetting "WINDOW", "left", Me.Left
        .SetSectionSetting "WINDOW", "top", Me.Top
        .SetSectionSetting "WINDOW", "height", Me.Height
        .SetSectionSetting "WINDOW", "width", Me.Width
        'Check size of grid
        lngColCount = graReport.DataGrid.ColumnCount
        lngRowcount = graReport.DataGrid.RowCount
        'Add column label info
        .AddSection "COLUMNLABELS"
        .SetSectionSetting "COLUMNLABELS", "columncount", lngColCount
        strDataLine = ""
        For lngColCounter = 1 To lngColCount
            strDataLine = strDataLine & graReport.DataGrid.ColumnLabel(lngColCounter, 1) & conSeperator
        Next
        strDataLine = Left$(strDataLine, Len(strDataLine) - 1) & ""
        .SetSectionSetting "COLUMNLABELS", "columnlabels", strDataLine
        'Add row label info
        .AddSection "ROWLABELS"
        .SetSectionSetting "ROWLABELS", "rowcount", lngRowcount
        strDataLine = ""
        For lngRowCounter = 1 To lngRowcount
            strDataLine = strDataLine & graReport.DataGrid.RowLabel(lngRowCounter, 1) & conSeperator
        Next
        strDataLine = Left$(strDataLine, Len(strDataLine) - 1) & ""
        .SetSectionSetting "ROWLABELS", "rowlabels", strDataLine
        'Add datagrid info
        .AddSection "DATAGRID"
        .SetSectionSetting "DATAGRID", "columncount", lngColCount
        .SetSectionSetting "DATAGRID", "rowcount", lngRowcount
        For lngRowCounter = 1 To lngRowcount
            strDataLine = ""
            For lngColCounter = 1 To lngColCount
                Call graReport.DataGrid.GetData(lngRowCounter, lngColCounter, dblDataPoint, intNullFlag)
                If intNullFlag = -1 Then
                    strDataLine = strDataLine & "NULL" & conSeperator
                Else
                    strDataLine = strDataLine & dblDataPoint & conSeperator
                End If
            Next
            strDataLine = Left$(strDataLine, Len(strDataLine) - 1) & ""
            .SetSectionSetting "DATAGRID", "row" & lngRowCounter, strDataLine
        Next
        .AddSection "CHART"
        .SetSectionSetting "CHART", "type", graReport.chartType
        .AddSection "PLOT"
        .SetSectionSetting "PLOT", "lrminx", graReport.Plot.LocationRect.Min.X
        .SetSectionSetting "PLOT", "lrminy", graReport.Plot.LocationRect.Min.Y
        .SetSectionSetting "PLOT", "lrmaxx", graReport.Plot.LocationRect.Max.X
        .SetSectionSetting "PLOT", "lrmaxy", graReport.Plot.LocationRect.Max.Y
        .AddSection "LEGEND"
        .SetSectionSetting "LEGEND", "lrminx", graReport.Legend.Location.Rect.Min.X
        .SetSectionSetting "LEGEND", "lrminy", graReport.Legend.Location.Rect.Min.Y
        .SetSectionSetting "LEGEND", "lrmaxx", graReport.Legend.Location.Rect.Max.X
        .SetSectionSetting "LEGEND", "lrmaxy", graReport.Legend.Location.Rect.Max.Y
        .FileSave strSaveFileName
    End With
    'Then set the caption of the form and the name of the data array
    Me.Caption = StripFileExtension(GetFileNameFromPath(strSaveFileName))
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
'Created by Robin G Brown, 19/5/97
'Print the graph
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
    'Stick the graphic on the clipboard
    Clipboard.Clear
    graReport.EditCopy
    'Print
    lngPicWidth = Clipboard.GetData(vbCFMetafile).Width
    lngPicHeight = Clipboard.GetData(vbCFMetafile).Height
    'Work out a scaling factor for displaying the grid
    sngScaleFactorActual = 1#
    With Printer
        'Try it portrait first
        .Orientation = vbPRORPortrait
        sngScaleFactorPortrait = 1#
        'It doesn't matter which is checked first, _
            height or width, as long as sngScaleFactorPortrait is  taken into account
        sngScaleFactorPortrait = _
            ((.Height - 800) / lngPicHeight) _
            * sngScaleFactorPortrait
        sngScaleFactorPortrait = _
            ((.Width - 600) / (lngPicWidth * sngScaleFactorPortrait)) _
            * sngScaleFactorPortrait
        'Then landscape
        .Orientation = vbPRORLandscape
        sngScaleFactorLandscape = 1#
        sngScaleFactorLandscape = _
            ((.Height - 800) / lngPicHeight) _
            * sngScaleFactorPortrait
        sngScaleFactorLandscape = _
            ((.Width - 600) / (lngPicWidth * sngScaleFactorPortrait)) _
            * sngScaleFactorPortrait
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
    Printer.PaintPicture Clipboard.GetData(vbCFMetafile), 0, 0, _
        Clipboard.GetData(vbCFMetafile).Width * sngScaleFactorActual, _
        Clipboard.GetData(vbCFMetafile).Height * sngScaleFactorActual
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

Public Sub GraphFileOpenMenuRoutine(ByVal strOpenFileName As String)
'Created by Robin G Brown, 19/5/97
'Retrieve the formatting and data from a file
'Sub Declares
Const conSub = "FileOpen"
Dim intFileNumber As Integer
Dim strDataLine As String
Dim intDataLineCounter As Integer
Dim strLineArray() As String
Dim lngSectionCounter As Long
Dim lngColCounter As Long
Dim lngRowCounter As Long
Dim lngColCount As Long
Dim lngRowcount As Long
Dim dblDataPoint As Double
Dim intNullFlag As Integer
    'Error Trap
    On Error GoTo GraphFileOpenMenuRoutine_ErrorHandler
    Set sffGraph = New SecFile
    strSaveFileName = strOpenFileName
    With sffGraph
        'Open the file
        .FileOpen strSaveFileName
        Me.Caption = StripFileExtension(GetFileNameFromPath(strSaveFileName))
        'Get the size of the grid first
        lngColCount = .GetSectionSetting("DATAGRID", "columncount")
        lngRowcount = .GetSectionSetting("DATAGRID", "rowcount")
        'Size the graph appropriately
        Call graReport.DataGrid.SetSize(1, 1, lngRowcount, lngColCount)
        For lngSectionCounter = 1 To .Sections
            Select Case .IndexSection(lngSectionCounter)
            Case "DATAGRID"
                ReDim strLineArray(1 To lngColCount)
                For lngRowCounter = 1 To lngRowcount
                    strDataLine = .GetSectionSetting("DATAGRID", "row" & lngRowCounter)
                    'Read CSV line into array
                    lngReturn = WriteCSVToArray(strDataLine, strLineArray())
                    For lngColCounter = 1 To lngColCount
                        If strLineArray(lngColCounter) = "NULL" Then
                            Call graReport.DataGrid.SetData(lngRowCounter, lngColCounter, 0, True)
                        Else
                            Call graReport.DataGrid.SetData(lngRowCounter, lngColCounter, CDbl(strLineArray(lngColCounter)), False)
                        End If
                    Next
                Next
            Case "ROWLABELS"
                strDataLine = .GetSectionSetting("ROWLABELS", "rowlabels")
                ReDim strLineArray(1 To CharactersInString(strDataLine, conSeperator) + 1)
                lngReturn = WriteCSVToArray(strDataLine, strLineArray())
                For lngColCounter = 1 To UBound(strLineArray, 1)
                    graReport.DataGrid.RowLabel(lngColCounter, 1) = strLineArray(lngColCounter)
                Next
            Case "COLUMNLABELS"
                strDataLine = .GetSectionSetting("COLUMNLABELS", "columnlabels")
                ReDim strLineArray(1 To CharactersInString(strDataLine, conSeperator) + 1)
                lngReturn = WriteCSVToArray(strDataLine, strLineArray())
                For lngColCounter = 1 To UBound(strLineArray, 1)
                    graReport.DataGrid.ColumnLabel(lngColCounter, 1) = strLineArray(lngColCounter)
                Next
            Case "WINDOW"
                Me.Left = .GetSectionSetting("WINDOW", "left")
                Me.Top = .GetSectionSetting("WINDOW", "top")
                Me.Height = .GetSectionSetting("WINDOW", "height")
                Me.Width = .GetSectionSetting("WINDOW", "width")
                Me.WindowState = .GetSectionSetting("WINDOW", "state")
            Case "CHART"
                graReport.chartType = .GetSectionSetting("CHART", "type")
            Case "PLOT"
                graReport.Plot.LocationRect.Min.X = .GetSectionSetting("PLOT", "lrminx")
                graReport.Plot.LocationRect.Min.Y = .GetSectionSetting("PLOT", "lrminy")
                graReport.Plot.LocationRect.Max.X = .GetSectionSetting("PLOT", "lrmaxx")
                graReport.Plot.LocationRect.Max.Y = .GetSectionSetting("PLOT", "lrmaxy")
            Case "LEGEND"
                graReport.Legend.Location.Rect.Min.X = .GetSectionSetting("LEGEND", "lrminx")
                graReport.Legend.Location.Rect.Min.Y = .GetSectionSetting("LEGEND", "lrminy")
                graReport.Legend.Location.Rect.Max.X = .GetSectionSetting("LEGEND", "lrmaxx")
                graReport.Legend.Location.Rect.Max.Y = .GetSectionSetting("LEGEND", "lrmaxy")
                graReport.Legend.Location.LocationType = VtChLocationTypeCustom
            Case Else
                'Ignore
            End Select
        Next
    End With
    'Tidy up the chart
    Call ColouriseChart
    Call ShowLegend(Me.graReport)
    graReport.Refresh
    booIsDirty = False
    booChartNew = False
Exit Sub
'Error Handler
GraphFileOpenMenuRoutine_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbOKOnly, conModule & ":" & conSub
    End Select
    Close #intFileNumber
    Exit Sub
End Sub

Private Sub Form_Terminate()
'Created by Robin G Brown, 29/8/97
    'Error Trap
    On Error Resume Next
    Set sffGraph = Nothing
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

Public Sub SetNextPen(ByRef penCurrentPen As Pen, ByVal bytPenStyle As Byte, ByVal bytPenColour As Byte)
'Created by Robin G Brown, 20/8/97
'Change the pen to the given style of pen
'Sub Declares
Const conSub = "SetNextPen"
Dim intRed As Integer
Dim intBlue As Integer
Dim intGreen As Integer
    'Error Trap
    On Error Resume Next
    With penCurrentPen
        'Width needs to be at least 40(ish) to be not a hairline! - RGB/21/8/97
        .Width = 40
        Select Case bytPenStyle
        Case 1
            .Style = VtPenStyleSolid
        Case 2
            .Style = VtPenStyleDashed
        Case 3
            .Style = VtPenStyleDitted
        Case 4
            .Style = VtPenStyleDotted
        Case 5
            .Style = VtPenStyleDashDit
        Case 6
            .Style = VtPenStyleDashDitDit
        Case 7
            .Style = VtPenStyleDashDot
        Case 8
            .Style = VtPenStyleDashDotDot
        Case Else
            .Style = VtPenStyleSolid
            '.Width = 0
        End Select
        If booMonoColour = True Then
            Select Case bytPenColour
            Case 1
                intRed = RedFromRGB(vbBlack)
                intGreen = GreenFromRGB(vbBlack)
                intBlue = BlueFromRGB(vbBlack)
                .VtColor.Set intRed, intGreen, intBlue
            Case 2
                intRed = RedFromRGB(vbRed)
                intGreen = GreenFromRGB(vbRed)
                intBlue = BlueFromRGB(vbRed)
                .VtColor.Set intRed, intGreen, intBlue
            Case 3
                intRed = RedFromRGB(vbGreen)
                intGreen = GreenFromRGB(vbGreen)
                intBlue = BlueFromRGB(vbGreen)
                .VtColor.Set intRed, intGreen, intBlue
            Case 4
                intRed = RedFromRGB(vbYellow)
                intGreen = GreenFromRGB(vbYellow)
                intBlue = BlueFromRGB(vbYellow)
                .VtColor.Set intRed, intGreen, intBlue
            Case 5
                intRed = RedFromRGB(vbBlue)
                intGreen = GreenFromRGB(vbBlue)
                intBlue = BlueFromRGB(vbBlue)
                .VtColor.Set intRed, intGreen, intBlue
            Case 6
                intRed = RedFromRGB(vbMagenta)
                intGreen = GreenFromRGB(vbMagenta)
                intBlue = BlueFromRGB(vbMagenta)
                .VtColor.Set intRed, intGreen, intBlue
            Case 7
                intRed = RedFromRGB(vbCyan)
                intGreen = GreenFromRGB(vbCyan)
                intBlue = BlueFromRGB(vbCyan)
                .VtColor.Set intRed, intGreen, intBlue
            Case 8
                intRed = RedFromRGB(vbWhite)
                intGreen = GreenFromRGB(vbWhite)
                intBlue = BlueFromRGB(vbWhite)
                .VtColor.Set intRed, intGreen, intBlue
            Case Else
                .VtColor.Set 0, 0, 0 'Black
            End Select
        Else
            .VtColor.Set 0, 0, 0 'Black
        End If
    End With
End Sub

Public Sub SetNextBrush(ByRef bshCurrentBrush As Brush, ByVal bytBrushColour As Byte)
'Created by Robin G Brown, 20/8/97
'Change the brush to the given style of brush
'Sub Declares
Const conSub = "SetNextBrush"
Dim intRed As Integer
Dim intBlue As Integer
Dim intGreen As Integer
    'Error Trap
    On Error Resume Next
    With bshCurrentBrush
        'Always use a solid brush
        .Style = VtBrushStyleSolid
        If booMonoColour = True Then
            Select Case bytBrushColour
            Case 0
                intRed = RedFromRGB(vbBlack)
                intGreen = GreenFromRGB(vbBlack)
                intBlue = BlueFromRGB(vbBlack)
                .FillColor.Set intRed, intGreen, intBlue
            Case 1
                intRed = RedFromRGB(vbRed)
                intGreen = GreenFromRGB(vbRed)
                intBlue = BlueFromRGB(vbRed)
                .FillColor.Set intRed, intGreen, intBlue
            Case 2
                intRed = RedFromRGB(vbGreen)
                intGreen = GreenFromRGB(vbGreen)
                intBlue = BlueFromRGB(vbGreen)
                .FillColor.Set intRed, intGreen, intBlue
            Case 3
                intRed = RedFromRGB(vbYellow)
                intGreen = GreenFromRGB(vbYellow)
                intBlue = BlueFromRGB(vbYellow)
                .FillColor.Set intRed, intGreen, intBlue
            Case 4
                intRed = RedFromRGB(vbBlue)
                intGreen = GreenFromRGB(vbBlue)
                intBlue = BlueFromRGB(vbBlue)
                .FillColor.Set intRed, intGreen, intBlue
            Case 5
                intRed = RedFromRGB(vbMagenta)
                intGreen = GreenFromRGB(vbMagenta)
                intBlue = BlueFromRGB(vbMagenta)
                .FillColor.Set intRed, intGreen, intBlue
            Case 6
                intRed = RedFromRGB(vbCyan)
                intGreen = GreenFromRGB(vbCyan)
                intBlue = BlueFromRGB(vbCyan)
                .FillColor.Set intRed, intGreen, intBlue
            Case 7
                intRed = RedFromRGB(vbWhite)
                intGreen = GreenFromRGB(vbWhite)
                intBlue = BlueFromRGB(vbWhite)
                .FillColor.Set intRed, intGreen, intBlue
            Case Else
                .FillColor.Set 0, 0, 0 'Black
            End Select
        Else
            .FillColor.Set 0, 0, 0 'Black
        End If
    End With
End Sub

