Attribute VB_Name = "modChartLib"
'Created by Robin G Brown, 6th May 1997
'-----------------------------------------------------------------------------
'   This module contains code for _
    handling chart controls
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Type SeriesInfo
    'Use sei as a prefix
    SeriesColumn As Long
    SeriesType As Integer
    SeriesColour As Long
    SeriesWidth As Byte
End Type
Private Const conModule = "modChartLib"
Private lngCounter As Long
Private lngReturn As Long
'An array to hold info about the series of the current chart we are working with
'Private seiSeriesInfo() As SeriesInfo
'Legend Alignment Constants
Public Const conAlignLegendTopCentre = 101
Public Const conAlignLegendTopRight = 102
Public Const conAlignLegendCentreRight = 103
Public Const conAlignLegendBottomRight = 104
Public Const conAlignLegendBottomCentre = 105
Public Const conAlignLegendBottomLeft = 106
Public Const conAlignLegendCentreLeft = 107
Public Const conAlignLegendTopLeft = 108
Public Const conAlignLegendCentreCentre = 109
'-----------------------------------------------------------------------------
'Spoiler information for charts
'
'   1. The Datagrid has an LBound of 1!
'   2. The datagrid matrix is COLS, ROWS - and so is SetData (COLS,ROWS)!
'   3. ColumnLabels are up the side (Y axis)
'   4. RowLabels are along the bottom (X axis)
'   5. The origin point for the plot is 0,0 = bottom left corner?
'-----------------------------------------------------------------------------

Public Sub RemoveColumn(ByRef graControl As MSChart, ByVal lngColumn As Long)
'Created by Robin G Brown, 18/9/97
'Set the selected column on the datagrid to zero and remove the labels
'Sub Declares
Dim lngRowCounter As Long
Dim lngColCounter As Long
Dim lngCols As Long
Dim lngRows As Long
Dim strColLabels() As String
Dim dblValues() As Double
Dim intNulls() As Integer
    'Error Trap
    On Error Resume Next
    'Copy the data to the arrays, missing out lngcolumn
    lngRows = graControl.DataGrid.RowCount
    ReDim dblValues(1 To lngRows, 1 To 1)
    ReDim intNulls(1 To lngRows, 1 To 1)
    ReDim strColLabels(1 To 1)
    lngCols = 0
    For lngColCounter = 1 To graControl.DataGrid.ColumnCount
        If lngColCounter <> lngColumn Then
            lngCols = lngCols + 1
            For lngRowCounter = 1 To lngRows
                ReDim Preserve strColLabels(1 To lngCols)
                strColLabels(lngCols) = _
                    graControl.DataGrid.ColumnLabel(lngColCounter, 1)
                ReDim Preserve dblValues(1 To lngRows, 1 To lngCols)
                ReDim Preserve intNulls(1 To lngRows, 1 To lngCols)
                Call graControl.DataGrid.GetData( _
                    lngRowCounter, _
                    lngColCounter, _
                    dblValues(lngRowCounter, lngCols), _
                    intNulls(lngRowCounter, lngCols))
            Next
        End If
    Next
    'Resize
    Call graControl.DataGrid.DeleteColumns(1, graControl.DataGrid.ColumnCount)
    Call graControl.DataGrid.SetSize(1, 1, lngRows, lngCols)
    'Write the data to the datagrid
    For lngColCounter = 1 To lngCols
        For lngRowCounter = 1 To lngRows
            graControl.DataGrid.ColumnLabel(lngColCounter, 1) = strColLabels(lngColCounter)
            Call graControl.DataGrid.SetData( _
                lngRowCounter, _
                lngColCounter, _
                dblValues(lngRowCounter, lngColCounter), _
                intNulls(lngRowCounter, lngColCounter))
        Next
    Next
End Sub

Public Sub RemoveAllColumns(ByRef graControl As MSChart)
'Created by Robin G Brown, 21/8/97
'Clear the chart of any information
    'Error Trap
    On Error Resume Next
    'Delete all the columns
    Call graControl.DataGrid.DeleteColumns(1, graControl.DataGrid.ColumnCount)
    'Resize the datagrid
    Call graControl.DataGrid.SetSize(1, 1, 1, 1)
End Sub

Public Sub DrawBlankAxes(ByRef graControl As MSChart)
'Created by Robin G Brown, 17/9/97
'Draw a blank chart with axes but no data
'Sub Declares
Dim intRowCounter As Integer
Dim intColCounter As Integer
    'Error Trap
    On Error Resume Next
    'Set datagrid
    Call graControl.DataGrid.SetSize(1, 1, 10, 100)
    'Fill in labels for rows
    For intRowCounter = 1 To 10
        graControl.DataGrid.RowLabel(intRowCounter, 1) = "" 'intRowCounter
    Next
    'Fill in label for column
    For intColCounter = 1 To 100
        graControl.DataGrid.ColumnLabel(1, 1) = "" 'intColCounter
    Next
    'Fill in data values with 0
    For intRowCounter = 1 To 10
        For intColCounter = 1 To 100
            Call graControl.DataGrid.SetData(intRowCounter, intColCounter, 0, True)
        Next
    Next
    graControl.Legend.Location.Visible = False
    'graControl.Plot.Axis(VtChAxisIdX).ValueScale.Maximum = 100
    'graControl.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 100
End Sub

Public Sub PlainChart(ByRef graControl As MSChart)
'Created by Robin G Brown, 13/5/97
'Make all the chart series very plain
'Sub Declares
    'Error Trap
    On Error Resume Next
    If graControl.chartType = VtChChartType2dLine _
    Or graControl.chartType = VtChChartType3dLine _
    Then
        For lngCounter = 1 To graControl.Plot.SeriesCollection.Count
            With graControl.Plot.SeriesCollection(lngCounter).Pen
                .VtColor.Set 0, 0, 0
                .Width = 0
            End With
        Next
    Else
        For lngCounter = 1 To graControl.Plot.SeriesCollection.Count
            With graControl.Plot.SeriesCollection(lngCounter)
                .DataPoints(-1).Brush.FillColor.Set 0, 0, 0
                .Pen.Width = 0
            End With
        Next
    End If
End Sub

Public Sub SetColumnColour(ByRef graControl As MSChart, ByVal intColumn As Integer, ByVal lngColour As Long)
'Created by Robin G Brown, 13/5/97
'Colour in the given column
'Sub Declares
Dim intRed As Integer
Dim intGreen As Integer
Dim intBlue As Integer
    'Error Trap
    On Error Resume Next
    intRed = RedFromRGB(lngColour)
    intGreen = GreenFromRGB(lngColour)
    intBlue = BlueFromRGB(lngColour)
    If graControl.chartType = VtChChartType2dLine _
    Or graControl.chartType = VtChChartType3dLine _
    Then
        With graControl.Plot.SeriesCollection(intColumn).Pen
            .VtColor.Set intRed, intGreen, intBlue
            '.Width = 1
        End With
    Else
        With graControl.Plot.SeriesCollection(intColumn)
            .DataPoints(-1).Brush.FillColor.Set intRed, intGreen, intBlue
            '.Pen.width = 1
        End With
    End If
End Sub

Public Sub MaximizePlot(ByRef graControl As MSChart)
'Created by Robin G Brown, 9/5/97
'Make the chart as big as possible within the bounds of the container with a healthy margin
'Sub Declares
    'Error Trap
    On Error Resume Next
    Call graControl.Plot.LocationRect.Min.Set(120, 120)
    Call graControl.Plot.LocationRect.Max.Set(graControl.Width - 120, graControl.Height - 120)
End Sub

Public Sub PositionLegendAndPlot(ByRef graControl As MSChart, ByVal intAlign As Integer, ByVal intPercentageWidth As Integer, ByVal intPercentageHeight As Integer)
'Created by Robin G Brown, 6/5/97
'Align the legend and maximize the plot in the leftover space
'Sub Declares
Dim sngChartHeight As Integer
Dim sngChartWidth As Integer
Dim sngLegendHeight As Integer
Dim sngLegendWidth As Integer
Dim sngHalfChartHeight As Integer
Dim sngHalfChartWidth As Integer
Dim sngHalfLegendHeight As Integer
Dim sngHalfLegendWidth As Integer
    'Error Trap
    On Error Resume Next
    'Set up a rect object intPercentageWidth% the width and _
        intPercentageHeight% the height of the control, _
        in the intAlign part of the control
    sngChartHeight = graControl.Height
    sngChartWidth = graControl.Width
    sngLegendHeight = sngChartHeight * (intPercentageHeight / 100)
    sngLegendWidth = sngChartWidth * (intPercentageWidth / 100)
    sngHalfChartHeight = sngChartHeight / 2
    sngHalfChartWidth = sngChartWidth / 2
    sngHalfLegendHeight = sngLegendHeight / 2
    sngHalfLegendWidth = sngLegendWidth / 2
    graControl.Legend.Location.LocationType = VtChLocationTypeCustom
    Select Case intAlign
    Case conAlignLegendTopCentre
        Call graControl.Legend.Location.Rect.Min.Set(sngHalfChartWidth - sngHalfLegendWidth, sngChartHeight - sngLegendHeight)
        Call graControl.Legend.Location.Rect.Max.Set(sngHalfChartWidth + sngHalfLegendWidth, sngChartHeight)
        Call graControl.Plot.LocationRect.Min.Set(120, 120)
        Call graControl.Plot.LocationRect.Max.Set(graControl.Width - 120, graControl.Height - sngLegendHeight - 120)
    Case conAlignLegendTopRight
        Call graControl.Legend.Location.Rect.Min.Set(sngChartWidth - sngLegendWidth, sngChartHeight - sngLegendHeight)
        Call graControl.Legend.Location.Rect.Max.Set(sngChartWidth, sngChartHeight)
        Call graControl.Plot.LocationRect.Min.Set(120, 120)
        Call graControl.Plot.LocationRect.Max.Set(graControl.Width - 120, graControl.Height - sngLegendHeight - 120)
    Case conAlignLegendCentreRight
        Call graControl.Legend.Location.Rect.Min.Set(sngChartWidth - sngLegendWidth, sngHalfChartHeight - sngHalfLegendHeight)
        Call graControl.Legend.Location.Rect.Max.Set(sngChartWidth, sngHalfChartHeight + sngHalfLegendHeight)
        Call graControl.Plot.LocationRect.Min.Set(120, 120)
        Call graControl.Plot.LocationRect.Max.Set(graControl.Width - sngLegendWidth - 120, graControl.Height - 120)
    Case conAlignLegendBottomRight
        Call graControl.Legend.Location.Rect.Min.Set(sngChartWidth - sngLegendWidth, 0)
        Call graControl.Legend.Location.Rect.Max.Set(sngChartWidth, sngLegendHeight)
        Call graControl.Plot.LocationRect.Min.Set(120, sngLegendHeight + 120)
        Call graControl.Plot.LocationRect.Max.Set(graControl.Width - 120, graControl.Height - 120)
    Case conAlignLegendBottomCentre
        Call graControl.Legend.Location.Rect.Min.Set(sngHalfChartWidth - sngHalfLegendWidth, 0)
        Call graControl.Legend.Location.Rect.Max.Set(sngHalfChartWidth + sngHalfLegendWidth, sngLegendHeight)
        Call graControl.Plot.LocationRect.Min.Set(120, sngLegendHeight + 120)
        Call graControl.Plot.LocationRect.Max.Set(graControl.Width - 120, graControl.Height - 120)
    Case conAlignLegendBottomLeft
        Call graControl.Legend.Location.Rect.Min.Set(0, 0)
        Call graControl.Legend.Location.Rect.Max.Set(sngLegendWidth, sngLegendHeight)
        Call graControl.Plot.LocationRect.Min.Set(120, sngLegendHeight + 120)
        Call graControl.Plot.LocationRect.Max.Set(graControl.Width - 120, graControl.Height - 120)
    Case conAlignLegendCentreLeft
        Call graControl.Legend.Location.Rect.Min.Set(0, sngHalfChartHeight - sngHalfLegendHeight)
        Call graControl.Legend.Location.Rect.Max.Set(sngHalfChartWidth - sngHalfLegendWidth, sngHalfChartHeight + sngHalfLegendHeight)
        Call graControl.Plot.LocationRect.Min.Set(sngLegendWidth + 120, 120)
        Call graControl.Plot.LocationRect.Max.Set(graControl.Width - (sngHalfChartWidth - sngHalfLegendWidth) - 120, graControl.Height - 120)
    Case conAlignLegendTopLeft
        Call graControl.Legend.Location.Rect.Min.Set(0, sngChartHeight - sngLegendHeight)
        Call graControl.Legend.Location.Rect.Max.Set(sngLegendWidth, sngChartHeight)
        Call graControl.Plot.LocationRect.Min.Set(120, 120)
        Call graControl.Plot.LocationRect.Max.Set(graControl.Width - 120, graControl.Height - sngLegendHeight - 120)
    Case conAlignLegendCentreCentre
        Call graControl.Legend.Location.Rect.Min.Set(sngHalfChartWidth - sngHalfLegendWidth, sngHalfChartHeight - sngHalfLegendHeight)
        Call graControl.Legend.Location.Rect.Max.Set(sngHalfChartWidth + sngHalfLegendWidth, sngHalfChartHeight + sngHalfLegendHeight)
        Call graControl.Plot.LocationRect.Min.Set(120, 120)
        Call graControl.Plot.LocationRect.Max.Set(graControl.Width - 120, graControl.Height - 120)
    End Select
End Sub

Public Sub PositionLegend(ByRef graControl As MSChart, ByVal intAlign As Integer, ByVal intPercentageWidth As Integer, ByVal intPercentageHeight As Integer)
'Created by Robin G Brown, 6/5/97
'Align the legend
'Sub Declares
Dim sngChartHeight As Integer
Dim sngChartWidth As Integer
Dim sngLegendHeight As Integer
Dim sngLegendWidth As Integer
Dim sngHalfChartHeight As Integer
Dim sngHalfChartWidth As Integer
Dim sngHalfLegendHeight As Integer
Dim sngHalfLegendWidth As Integer
    'Error Trap
    On Error Resume Next
    'Set up a rect object intPercentageWidth% the width and _
        intPercentageHeight% the height of the control, _
        in the intAlign part of the control
    sngChartHeight = graControl.Height
    sngChartWidth = graControl.Width
    sngLegendHeight = sngChartHeight * (intPercentageHeight / 100)
    sngLegendWidth = sngChartWidth * (intPercentageWidth / 100)
    sngHalfChartHeight = sngChartHeight / 2
    sngHalfChartWidth = sngChartWidth / 2
    sngHalfLegendHeight = sngLegendHeight / 2
    sngHalfLegendWidth = sngLegendWidth / 2
    Select Case intAlign
    Case conAlignLegendTopCentre
        Call graControl.Legend.Location.Rect.Min.Set(sngHalfChartWidth - sngHalfLegendWidth, sngChartHeight - sngLegendHeight)
        Call graControl.Legend.Location.Rect.Max.Set(sngHalfChartWidth + sngHalfLegendWidth, sngChartHeight)
    Case conAlignLegendTopRight
        Call graControl.Legend.Location.Rect.Min.Set(sngChartWidth - sngLegendWidth, sngChartHeight - sngLegendHeight)
        Call graControl.Legend.Location.Rect.Max.Set(sngChartWidth, sngChartHeight)
    Case conAlignLegendCentreRight
        Call graControl.Legend.Location.Rect.Min.Set(sngChartWidth - sngLegendWidth, sngHalfChartHeight - sngHalfLegendHeight)
        Call graControl.Legend.Location.Rect.Max.Set(sngChartWidth, sngHalfChartHeight + sngHalfLegendHeight)
    Case conAlignLegendBottomRight
        Call graControl.Legend.Location.Rect.Min.Set(sngChartWidth - sngLegendWidth, 0)
        Call graControl.Legend.Location.Rect.Max.Set(sngChartWidth, sngLegendHeight)
    Case conAlignLegendBottomCentre
        Call graControl.Legend.Location.Rect.Min.Set(sngHalfChartWidth - sngHalfLegendWidth, 0)
        Call graControl.Legend.Location.Rect.Max.Set(sngHalfChartWidth + sngHalfLegendWidth, sngLegendHeight)
    Case conAlignLegendBottomLeft
        Call graControl.Legend.Location.Rect.Min.Set(0, 0)
        Call graControl.Legend.Location.Rect.Max.Set(sngLegendWidth, sngLegendHeight)
    Case conAlignLegendCentreLeft
        Call graControl.Legend.Location.Rect.Min.Set(0, sngHalfChartHeight - sngHalfLegendHeight)
        Call graControl.Legend.Location.Rect.Max.Set(sngHalfChartWidth - sngHalfLegendWidth, sngHalfChartHeight + sngHalfLegendHeight)
    Case conAlignLegendTopLeft
        Call graControl.Legend.Location.Rect.Min.Set(0, sngChartHeight - sngLegendHeight)
        Call graControl.Legend.Location.Rect.Max.Set(sngLegendWidth, sngChartHeight)
    Case conAlignLegendCentreCentre
        Call graControl.Legend.Location.Rect.Min.Set(sngHalfChartWidth - sngHalfLegendWidth, sngHalfChartHeight - sngHalfLegendHeight)
        Call graControl.Legend.Location.Rect.Max.Set(sngHalfChartWidth + sngHalfLegendWidth, sngHalfChartHeight + sngHalfLegendHeight)
    End Select
    graControl.Legend.Location.LocationType = VtChLocationTypeCustom
End Sub

Public Sub ShowLegend(ByRef graControl As MSChart)
'Created by Robin G Brown, 9/5/97
'Shows the legend
'Sub Declares
    'Error Trap
    On Error Resume Next
    graControl.Legend.Location.Visible = True
End Sub

Public Sub HideLegend(ByRef graControl As MSChart)
'Created by Robin G Brown, 9/5/97
'Hides the legend
'Sub Declares
    'Error Trap
    On Error Resume Next
    graControl.Legend.Location.Visible = False
End Sub

'-----------------------------------------------------------------------------
' These three functions are copied from the
'   Manipulating the Chart Appearance' topic in MSCHART.HLP
'-----------------------------------------------------------------------------
Public Function RedFromRGB(ByVal lngRGBValue As Long) As Integer
    ' The ampersand after &HFF coerces the number as a
    ' long, preventing Visual Basic from evaluating the
    ' number as a negative value. The logical And is
    ' used to return bit values.
    On Error Resume Next
    RedFromRGB = &HFF& And lngRGBValue
End Function

Public Function GreenFromRGB(ByVal lngRGBValue As Long) As Integer
    ' The result of the And operation is divided by
    ' 256, to return the value of the middle bytes.
    ' Note the use of the Integer divisor.
    On Error Resume Next
    GreenFromRGB = (&HFF00& And lngRGBValue) \ 256
End Function

Public Function BlueFromRGB(ByVal lngRGBValue As Long) As Integer
    ' This function works like the GreenFromlngRGBValue above,
    ' except you don't need the ampersand. The
    ' number is already a long. The result divided by
    ' 65536 to obtain the highest bytes.
    On Error Resume Next
    BlueFromRGB = (&HFF0000 And lngRGBValue) \ 65536
End Function
'-----------------------------------------------------------------------------

Public Sub GenericFillChartFromArray(ByRef graControl As MSChart, ByRef varArray As Variant)
'Created by Robin G Brown, 6/5/97
'This will need reworking for individual purposes...
'Set up a chart from an arraty, _
    assuming that the array matrix is (rows,cols) _
    and that both dimensions start with 0 _
    and that the first(0) col and row contain headers/labels
'Sub Declares
Const conSub = "GenericFillChartFromArray"
Dim intRowCounter As Integer
Dim intColCounter As Integer
Dim intCols As Integer
Dim intRows As Integer
    'Error Trap
    On Error GoTo GenericFillChartFromArray_ErrorHandler
    'Set the size of the datagrid first
    intRows = UBound(varArray, 1)
    intCols = UBound(varArray, 2)
    Call graControl.DataGrid.SetSize(1, 1, intRows, intCols)
    'Fill in labels for rows
    For intRowCounter = 1 To intRows
        graControl.DataGrid.RowLabel(intRowCounter, 1) = _
            CStr(varArray(intRowCounter, 0))
    Next
    'Fill in labels for columns
    For intColCounter = 1 To intCols
        graControl.DataGrid.ColumnLabel(intColCounter, 1) = _
            CStr(varArray(0, intColCounter))
    Next
    'Fill in data values
    For intRowCounter = 1 To intRows
        For intColCounter = 1 To intCols
            If IsNumeric(varArray(intRowCounter, intColCounter)) = True Then
                Call graControl.DataGrid.SetData _
                    (intRowCounter, intColCounter, _
                    CDbl(varArray(intRowCounter, intColCounter)), False)
            Else
                Call graControl.DataGrid.SetData _
                    (intRowCounter, intColCounter, _
                    0#, True)
            End If
        Next
    Next
Exit Sub
'Error Handler
GenericFillChartFromArray_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub



