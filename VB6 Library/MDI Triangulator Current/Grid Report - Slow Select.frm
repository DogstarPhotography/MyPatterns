VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmGridReport 
   Caption         =   "Report"
   ClientHeight    =   5475
   ClientLeft      =   1155
   ClientTop       =   1560
   ClientWidth     =   9105
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
   Icon            =   "Grid Report - Slow Select.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5475
   ScaleWidth      =   9105
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1020
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin MSFlexGridLib.MSFlexGrid grdReport 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   8070
      _Version        =   327680
      BackColor       =   8421504
      GridColor       =   4210752
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   0
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
    showing a report grid
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "frmBoilerPlate"
Private intCounter As Integer
Private intReturn As Integer
Private intColCounter As Integer
Private intRowCounter As Integer
Public dtaDataArray As TriDataArray
Private fccOriginalColours() As FlexGridCellColour
Private intShiftState As Integer

Public Sub DrawGrid()
'Created by Robin G Brown, 14/5/97
'Redisplay the grid with the current settings
'Sub Declares
    'Error Trap
    On Error Resume Next
    'Fill the grid from dtaDataArray
    grdReport.Redraw = False
    Call dtaDataArray.FillGrid(grdReport)
    Call AutoSizeGrid(grdReport)
    Call AlignAllText(grdReport, conAlignTextCenterCenter)
    Call ReadOriginalGridColours
    Call ColourBottomRight
    grdReport.Redraw = True
End Sub

Private Sub Form_Load()
'Created by Robin G Brown, 2/5/97
'Default Behaviour
'Sub Declares
Const conSub = "Form_Load"
    'Error Trap
    On Error GoTo Form_Load_ErrorHandler
    'Make sure the datagrid array is set up
    Set dtaDataArray = New TriDataArray
    Call dtaDataArray.CreateData
    dtaDataArray.Name = "Report"
    Me.Caption = dtaDataArray.Name
    'Fill the grid from dtaDataArray
    Call DrawGrid
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
    'Resize the grid to fill the form
    grdReport.Height = Me.Height - 360
    grdReport.Width = Me.Width - 120
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

Private Sub grdReport_Click()
'Created by Robin G Brown, 14/5/97
'Select items depending on where the click took place
'Sub Declares
Const conSub = "grdReport_Click"
Dim intAccount As Integer
    'Error Trap
    On Error GoTo grdReport_Click_ErrorHandler
    'MsgBox fgcCurrent.GridRow & ":" & fgcCurrent.GridCol
    Select Case fgcCurrent.GridRow
    Case 0, 1
        If fgcCurrent.GridCol < grdReport.FixedRows Then
            'Clear selection
            Call SelectGridItems(conSelectClear)
        Else
            'Select this column
            If frmMaster.stbStatus.Panels("panYears").Tag = "ON" Then
                Call SelectGridItems(conSelectAccount)
            Else
                Call SelectGridItems(conSelectColumn)
            End If
        End If
    Case Else
        'Check for click in fixed cols
        If fgcCurrent.GridCol < grdReport.FixedRows Then
            'Select a period
            If frmMaster.stbStatus.Panels("panYears").Tag = "ON" Then
                Call SelectGridItems(conSelectPeriod)
            Else
                Call SelectGridItems(conSelectYear)
            End If
        Else
            'Clear selection
            Call SelectGridItems(conSelectCell)
        End If
    End Select
Exit Sub
'Error Handler
grdReport_Click_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Private Sub grdReport_DblClick()
'Created by Robin G Brown, 14/5/97
'Select items depending on where the dblclick took place
'Sub Declares
Const conSub = "grdReport_DblClick"
Dim intAccount As Integer
    'Error Trap
    On Error GoTo grdReport_DblClick_ErrorHandler
    'MsgBox fgcCurrent.GridRow & ":" & fgcCurrent.GridCol
    If fgcCurrent.GridRow >= grdReport.FixedRows _
    And fgcCurrent.GridCol >= grdReport.FixedCols Then
        'Show this column (or account, if and other years is selected) _
            in a chart report form
        'Call SelectGridItems(conSelectColumn)
        If frmMaster.stbStatus.Panels("panYears").Tag = "OFF" Then
            'Single column
            Load frmChartReport
            intAccount = dtaDataArray.GetAccountFromColumn(fgcCurrent.GridCol)
            Call dtaDataArray.FillChart(frmChartReport.graReport, grdReport, intAccount, fgcCurrent.GridCol)
            Call PositionLegend(frmChartReport.graReport, conAlignLegendBottomCentre, 90, 10)
            Call MaximizePlot(frmChartReport.graReport)
            'Call PlainChart(frmChartReport.graReport)
            frmChartReport.Show
        Else
            'Account
            Load frmChartReport
            intAccount = dtaDataArray.GetAccountFromColumn(fgcCurrent.GridCol)
            Call dtaDataArray.FillChart(frmChartReport.graReport, grdReport, intAccount, 0)
            Call PositionLegend(frmChartReport.graReport, conAlignLegendBottomCentre, 90, 10)
            Call MaximizePlot(frmChartReport.graReport)
            'Call PlainChart(frmChartReport.graReport)
            frmChartReport.Show
        End If
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

Private Sub grdReport_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'Created by Robin G Brown, 13/5/97
'Make sure that fgcCurrent is set
'Sub Declares
    'Error Trap
    On Error Resume Next
    'First we need to find out where the mouse is
    Call GetCellFromXY(grdReport, x, y, fgcCurrent)
    'And what keys were pressed at the time
    intShiftState = Shift
    'Then popup a menu if required
    If Button = vbRightButton Then
        'CODE_HERE
    End If
End Sub
'Private Sub grdReport_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
''Created by Robin G Brown, 13/5/97
''Make sure that fgcCurrent is set
''Sub Declares
'    'Error Trap
'    On Error Resume Next
'    'First we need to find out where the mouse is
'    Call GetCellFromXY(grdReport, x, y, fgcCurrent)
'    'And what keys were pressed at the time
'    intShiftState = Shift
'End Sub

Public Sub ReadOriginalGridColours()
'Created by Robin G Brown, 16/5/97
'Set the colours
'Sub Declares
    'Error Trap
    On Error Resume Next
    With grdReport
        ReDim fccOriginalColours(1 To .Cols - 1)
        .Row = 2
        For intColCounter = 1 To .Cols - 1
            .Col = intColCounter
            fccOriginalColours(intColCounter).ForeColor = .CellForeColor
            fccOriginalColours(intColCounter).BackColor = .CellBackColor
        Next
    End With
End Sub

Public Sub SelectGridItems(ByVal intInstruction As Integer)
'Created by Robin G Brown, 13/5/97
'Select various elements of the grid by colouring them differently
'Sub Declares
Const conSub = "SelectGridItems"
Dim intStartCol As Integer
Dim intEndCol As Integer
Dim intCurrentRow As Integer
Dim strSearch As String
Dim intNumAccounts As Integer
Dim lngColour As Long
Dim intShiftDown As Integer
Dim intCtrlDown As Integer
Static intPreviousCol As Integer
Static intPreviousRow As Integer
Dim intSelectedRow As Integer
Dim intHiddenColumns As Integer
        'Error Trap
    On Error GoTo SelectGridItems_ErrorHandler
    With grdReport
        .Redraw = False
        'Find the first visible column
        For intColCounter = .FixedCols To .Cols - 1
            If .ColIsVisible(intColCounter) = True Then
                intHiddenColumns = intColCounter - .FixedCols
                Exit For
            End If
        Next
        If .TextMatrix(fgcCurrent.GridRow, intHiddenColumns + 2) = "" Then
            'No text in current cell so it may not be selected
            Beep
            .Redraw = True
            Exit Sub
        End If
        'Reset all columns according to the original
        .FillStyle = flexFillRepeat
        intShiftDown = (intShiftState And vbShiftMask) > 0
        intCtrlDown = (intShiftState And vbCtrlMask) > 0
        'Only clear if we are not multiselecting
        'If intShiftDown = True Or intCtrlDown = True Then
        If intCtrlDown = True Then
            'Do not clear
        Else
            'Clear
            For intColCounter = 1 To .Cols - 1
                .Col = intColCounter
                .Row = 2
                .RowSel = .Rows - 1
                .CellForeColor = fccOriginalColours(intColCounter).ForeColor
                .CellBackColor = fccOriginalColours(intColCounter).BackColor
            Next
        End If
        'Then Colour in cells according to fgcCurrent and intInstruction
        Select Case intInstruction
        Case conSelectClear
            'Do nothing
        Case conSelectColumn
            If intShiftDown = True Then
                .Col = intPreviousCol
            Else
                .Col = fgcCurrent.GridCol
            End If
            .Row = .FixedRows
            .ColSel = fgcCurrent.GridCol
            .RowSel = .Rows - 1
            .CellForeColor = .ForeColorSel
            .CellBackColor = .BackColorSel
        'Case conSelectRow
        '    .Col = .FixedCols
        '    .Row = fgcCurrent.GridRow
        '    .ColSel = .Cols - 1
        '    .RowSel = fgcCurrent.GridRow
        '    .CellForeColor = .ForeColorSel
        '    .CellBackColor = .BackColorSel
        Case conSelectCell
            .FillStyle = flexFillSingle
            .Col = fgcCurrent.GridCol
            .Row = fgcCurrent.GridRow
            .CellForeColor = .ForeColorSel
            .CellBackColor = .BackColorSel
        Case conSelectYear
            intStartCol = intHiddenColumns + 2
            strSearch = .TextMatrix(0, intStartCol)
            Do
                intStartCol = intStartCol - 1
            Loop While strSearch = .TextMatrix(0, intStartCol)
            intEndCol = intHiddenColumns + 2
            strSearch = .TextMatrix(0, intEndCol)
            Do
                intEndCol = intEndCol + 1
            Loop While strSearch = .TextMatrix(0, intEndCol)
            If intShiftDown = True Then
                .Col = intStartCol + 1
                .Row = fgcCurrent.GridRow
                .ColSel = intEndCol - 1
                .RowSel = intPreviousRow
                .CellForeColor = .ForeColorSel
                .CellBackColor = .BackColorSel
            Else
                .Col = intStartCol + 1
                .Row = fgcCurrent.GridRow
                .ColSel = intEndCol - 1
                .RowSel = fgcCurrent.GridRow
                .CellForeColor = .ForeColorSel
                .CellBackColor = .BackColorSel
            End If
        Case conSelectPeriod
            If intShiftDown = True Then
                'Select all periods between this period and the previous period
                'intSelectedRow = fgcCurrent.GridRow
                intNumAccounts = dtaDataArray.VisibleAccounts
                intSelectedRow = fgcCurrent.GridRow + ((intHiddenColumns \ intNumAccounts) * dtaDataArray.Periods)
                For intCurrentRow = iMin(intSelectedRow, intPreviousRow) To iMax(intSelectedRow, intPreviousRow)
                    For intCounter = 0 To dtaDataArray.UWYears - 1
                        If intCurrentRow - (intCounter * dtaDataArray.Periods) >= .FixedRows Then
                            .Col = intCounter * intNumAccounts + 2
                            .Row = intCurrentRow - (intCounter * dtaDataArray.Periods)
                            .ColSel = .Col + intNumAccounts - 1
                            '.RowSel = intCurrentRow - intCounter
                            .CellForeColor = .ForeColorSel
                            .CellBackColor = .BackColorSel
                        End If
                    Next
                Next
            Else
                'Select this period
                intNumAccounts = dtaDataArray.VisibleAccounts
                intCurrentRow = fgcCurrent.GridRow + ((intHiddenColumns \ intNumAccounts) * dtaDataArray.Periods)
                For intCounter = 0 To dtaDataArray.UWYears - 1
                    If intCurrentRow - (intCounter * dtaDataArray.Periods) >= .FixedRows Then
                        .Col = intCounter * intNumAccounts + 2
                        .Row = intCurrentRow - (intCounter * dtaDataArray.Periods)
                        .ColSel = .Col + intNumAccounts - 1
                        '.RowSel = intCurrentRow - intCounter
                        .CellForeColor = .ForeColorSel
                        .CellBackColor = .BackColorSel
                    End If
                Next
            End If
        Case conSelectAccount
            'Search for each column with a similar account name and select it
            If intShiftDown = True Then
                intStartCol = iMin(fgcCurrent.GridCol, intPreviousCol)
                intEndCol = iMax(fgcCurrent.GridCol, intPreviousCol)
                strSearch = .TextMatrix(1, intStartCol)
                For intCounter = 1 To .Cols - 1
                    If .TextMatrix(1, intCounter) = strSearch Then
                        .Col = intCounter
                        .Row = .FixedRows
                        .ColSel = intCounter + (intEndCol - intStartCol)
                        .RowSel = .Rows - 1
                        .CellForeColor = .ForeColorSel
                        .CellBackColor = .BackColorSel
                    End If
                Next
            Else
                strSearch = .TextMatrix(1, fgcCurrent.GridCol)
                For intCounter = 1 To .Cols - 1
                    If .TextMatrix(1, intCounter) = strSearch Then
                        .Col = intCounter
                        .ColSel = intCounter
                        .Row = .FixedRows
                        .RowSel = .Rows - 1
                        .CellForeColor = .ForeColorSel
                        .CellBackColor = .BackColorSel
                    End If
                Next
            End If
        Case Else
            'Do nothing
        End Select
        Call ColourBottomRight
        .Redraw = True
        .Col = fgcCurrent.GridCol
        .Row = fgcCurrent.GridRow
        intPreviousCol = fgcCurrent.GridCol
        intPreviousRow = fgcCurrent.GridRow
    End With
Exit Sub
'Error Handler
SelectGridItems_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Public Sub ColourBottomRight()
'Created by Robin G Brown, 23/5/97
'Colours in the bottom right set of cells
'Sub Declares
Const conSub = "ColourBottomRight"
Dim intNumVisibleAccounts As Integer
Dim intColOffset As Integer
    'Error Trap
    On Error GoTo ColourBottomRight_ErrorHandler
    With grdReport
        intColOffset = 2 + dtaDataArray.VisibleAccounts
        intRowCounter = .Rows - 1
        .FillStyle = flexFillRepeat
        For intRowCounter = .Rows - 1 To .FixedRows + dtaDataArray.Periods Step -1
            For intColCounter = intColOffset To .Cols - 1
                If .TextMatrix(intRowCounter, intColCounter) = "" Then
                    intColOffset = intColCounter
                    Exit For
                End If
            Next
            .Col = intColOffset
            .Row = intRowCounter
            .ColSel = .Cols - 1
            .CellBackColor = .BackColorBkg
        Next
    End With
Exit Sub
'Error Handler
ColourBottomRight_ErrorHandler:
    Exit Sub
End Sub

