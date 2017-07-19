Attribute VB_Name = "modFlexGridLib"
'Created by Robin G Brown, 2/5/97
'-----------------------------------------------------------------------------
'   This module contains code for _
    handling MSFlexGrid controls
'   Note that there may be code in this library that could be duplicated in _
    modGridLib - the Grid Control Library
'   Note that 'Mouse Library.bas' is needed to compile this library
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "modFlexGridLib"
Private intColCounter As Integer
Private intRowCounter As Integer
Private intReturn As Integer
'This type can be used to set _
    multiple data values in a function that interrogates a grid
Public Type FlexGridCell
    'Use fgc as a prefix
    GridCol As Long
    GridRow As Long
    GridText As String
    ForeColor As Long
    BackColor As Long
End Type
'Text alignment Constants
Public Const conAlignTextLeftTop = flexAlignLeftTop
Public Const conAlignTextLeftCenter = flexAlignLeftCenter
Public Const conAlignTextLeftBottom = flexAlignLeftBottom
Public Const conAlignTextCenterTop = flexAlignCenterTop
Public Const conAlignTextCenterCenter = flexAlignCenterCenter
Public Const conAlignTextCenterBottom = flexAlignCenterBottom
Public Const conAlignTextRightTop = flexAlignRightTop
Public Const conAlignTextRightCenter = flexAlignRightCenter
Public Const conAlignTextRightBottom = flexAlignRightBottom
Public Const conAlignTextGeneral = flexAlignGeneral
'Selection constants
Public Const conSelectClear = 0
Public Const conSelectColumn = 1
Public Const conSelectRow = 2
Public Const conSelectCell = 3
Public Const conSelectMulti = 16
'-----------------------------------------------------------------------------
'   Spoiler information for Grids
'
'   1   2   3   4   5   6   7   8   9
'   1   2   3   4   5   6   7   8   9
'   1   2   3   4   5   6   7   8   9
'   1   2   3   4   5   6   7   8   9
'
'-----------------------------------------------------------------------------

Public Function GetCellFromXY(ByVal grdControl As MSFlexGrid, sngXPos As Single, sngYPos As Single, ByRef fgcCell As FlexGridCell) As Boolean
'Created by Robin G Brown, 13/5/97
'Given an XY coord, return the row and col at XY
'Spoiler information...
'   0,0 ---->+X
'   |
'   |      COORDS=(X,Y)
'   V
'   +Y
'Sub Declares
Const conSub = "GetCellFromXY"
Dim intTotal As Integer
    'Error Trap
    On Error GoTo GetCellFromXY_ErrorHandler
    'Set up defaults first
    fgcCell.GridCol = -1
    fgcCell.GridRow = -1
    fgcCell.GridText = ""
    'Count the visible columns in grdControl _
        measuring the widths until we find a value > X _
        or we run out of columns
    'First check to see if it is in the fixed area
    intTotal = 0
    For intColCounter = 0 To grdControl.FixedCols - 1
        intTotal = intTotal + grdControl.ColWidth(intColCounter)
        If sngXPos <= intTotal Then
            'It's this column!
            fgcCell.GridCol = intColCounter
            Exit For
        End If
    Next
    'If we haven't found it yet it can't be in the fixed area
    If fgcCell.GridCol = -1 Then
        'Find the first visible column
        intColCounter = grdControl.FixedCols
        Do While grdControl.ColIsVisible(intColCounter) = False
            intColCounter = intColCounter + 1
        Loop
        'intColCounter now = first visible column
        Do While intColCounter < grdControl.Cols
            intTotal = intTotal + grdControl.ColWidth(intColCounter)
            If sngXPos <= intTotal Then
                'It's this column!
                fgcCell.GridCol = intColCounter
                Exit Do
            End If
            intColCounter = intColCounter + 1
        Loop
    End If
    'If we haven't found it it's not over a cell
    If fgcCell.GridCol = -1 Then
        'Haven't found it so...
        Err.Raise 55000
    End If
    'Count the visible rows in grdControl _
        measuring the heights until we find a value > y _
        or we run out of rows
    'First check to see if it is in the fixed area
    intTotal = 0
    For intRowCounter = 0 To grdControl.FixedRows - 1
        intTotal = intTotal + grdControl.RowHeight(intRowCounter)
        If sngYPos <= intTotal Then
            'It's this row!
            fgcCell.GridRow = intRowCounter
            Exit For
        End If
    Next
    'If we haven't found it yet it can't be in the fixed area
    If fgcCell.GridRow = -1 Then
        'Find the first visible row
        intRowCounter = grdControl.FixedRows
        Do While grdControl.RowIsVisible(intRowCounter) = False
            intRowCounter = intRowCounter + 1
        Loop
        'intRowCounter now = first visible row
        Do While intRowCounter < grdControl.Rows
            intTotal = intTotal + grdControl.RowHeight(intRowCounter)
            If sngYPos <= intTotal Then
                'It's this row!
                fgcCell.GridRow = intRowCounter
                Exit Do
            End If
            intRowCounter = intRowCounter + 1
        Loop
    End If
    'If we haven't found it it's not over a cell
    If fgcCell.GridRow = -1 Then
        'Haven't found it so...
        Err.Raise 55000
    End If
    fgcCell.GridText = grdControl.TextMatrix(fgcCell.GridRow, fgcCell.GridCol)
    GetCellFromXY = True
Exit Function
'Error Handler
GetCellFromXY_ErrorHandler:
    Select Case Err.Number
    Case 55000
        'Do nothing, no message required
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    fgcCell.GridCol = -1
    fgcCell.GridRow = -1
    fgcCell.GridText = ""
    GetCellFromXY = False
    Exit Function
End Function

Public Function GetCellFromMouse(ByVal grdControl As MSFlexGrid, ByRef mstGridMouse As MouseState, ByRef fgcCell As FlexGridCell) As Boolean
'Created by Robin G Brown, 13/5/97
'Given an XY coord, return the row and col at XY, using the MouseState type from the Mouse Library (modMouseLib)
'Spoiler information...
'   0,0 ---->+X
'   |
'   |      COORDS=(X,Y)
'   V
'   +Y
'  Note that no account is taken of the width of gridlines!
'Sub Declares
Const conSub = "GetCellFromMouse"
Dim intTotal As Integer
    'Error Trap
    On Error GoTo GetCellFromMouse_ErrorHandler
    'Set up defaults first
    fgcCell.GridCol = -1
    fgcCell.GridRow = -1
    fgcCell.GridText = ""
    'Count the visible columns in grdControl _
        measuring the widths until we find a value > X _
        or we run out of columns
    'First check to see if it is in the fixed area
    intTotal = 0
    For intColCounter = 0 To grdControl.FixedCols - 1
        intTotal = intTotal + grdControl.ColWidth(intColCounter)
        If mstGridMouse.x <= intTotal Then
            'It's this column!
            fgcCell.GridCol = intColCounter
            Exit For
        End If
    Next
    'If we haven't found it yet it can't be in the fixed area
    If fgcCell.GridCol = -1 Then
        'Find the first visible column
        intColCounter = grdControl.FixedCols
        Do While grdControl.ColIsVisible(intColCounter) = False
            intColCounter = intColCounter + 1
        Loop
        'intColCounter now = first visible column
        Do While intColCounter < grdControl.Cols
            intTotal = intTotal + grdControl.ColWidth(intColCounter)
            If mstGridMouse.x <= intTotal Then
                'It's this column!
                fgcCell.GridCol = intColCounter
                Exit Do
            End If
            intColCounter = intColCounter + 1
        Loop
    End If
    'If we haven't found it it's not over a cell
    If fgcCell.GridCol = -1 Then
        'Haven't found it so...
        Err.Raise 55000
    End If
    'Count the visible rows in grdControl _
        measuring the heights until we find a value > y _
        or we run out of rows
    'First check to see if it is in the fixed area
    intTotal = 0
    For intRowCounter = 0 To grdControl.FixedRows - 1
        intTotal = intTotal + grdControl.RowHeight(intRowCounter)
        If mstGridMouse.y <= intTotal Then
            'It's this row!
            fgcCell.GridRow = intRowCounter
            Exit For
        End If
    Next
    'If we haven't found it yet it can't be in the fixed area
    If fgcCell.GridRow = -1 Then
        'Find the first visible row
        intRowCounter = grdControl.FixedRows
        Do While grdControl.RowIsVisible(intRowCounter) = False
            intRowCounter = intRowCounter + 1
        Loop
        'intRowCounter now = first visible row
        Do While intRowCounter < grdControl.Rows
            intTotal = intTotal + grdControl.RowHeight(intRowCounter)
            If mstGridMouse.y <= intTotal Then
                'It's this row!
                fgcCell.GridRow = intRowCounter
                Exit Do
            End If
            intRowCounter = intRowCounter + 1
        Loop
    End If
    'If we haven't found it it's not over a cell
    If fgcCell.GridRow = -1 Then
        'Haven't found it so...
        Err.Raise 55000
    End If
    fgcCell.GridText = grdControl.TextMatrix(fgcCell.GridRow, fgcCell.GridCol)
    GetCellFromMouse = True
Exit Function
'Error Handler
GetCellFromMouse_ErrorHandler:
    Select Case Err.Number
    Case 55000
        'Do nothing, no message required
    Case Else
        'MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    fgcCell.GridCol = -1
    fgcCell.GridRow = -1
    fgcCell.GridText = ""
    GetCellFromMouse = False
    Exit Function
End Function

Public Sub AlignAllText(ByRef grdControl As MSFlexGrid, ByVal intAlign As Integer)
'Created by Robin G Brown, 7/5/97
'Align the text in the given column
'Sub Declares
    'Error Trap
    On Error Resume Next
    'For each column call AlignText
    For intColCounter = 0 To grdControl.Cols - 1
        grdControl.ColAlignment(intColCounter) = intAlign
    Next
End Sub

Public Sub AlignText(ByRef grdControl As MSFlexGrid, ByVal intColumn As Integer, ByVal intAlign As Integer)
'Created by Robin G Brown, 7/5/97
'Align the text in the given column
'Sub Declares
    'Error Trap
    On Error Resume Next
    'Align all the text in the given column according to intAlign
    grdControl.ColAlignment(intColumn) = intAlign
End Sub

Public Sub AutoSizeGrid(ByRef grdControl As MSFlexGrid)
'Created by Robin G Brown, 7/5/97
'Autosize all the columns
'Sub Declares
Const conSub = "AutoSizeGrid"
    'Error Trap
    On Error GoTo AutoSizeGrid_ErrorHandler
    'For each column call autosizecolumn
    For intColCounter = 0 To grdControl.Cols - 1
        Call AutoSizeColumn(grdControl, intColCounter)
    Next
Exit Sub
'Error Handler
AutoSizeGrid_ErrorHandler:
    Exit Sub
End Sub

Public Sub AutoSizeColumn(ByRef grdControl As MSFlexGrid, ByVal intColumn As Integer)
'Created by Robin G Brown, 7/5/97
'Autosize all the columns
'Sub Declares
Const conSub = "AutoSizeColumn"
Dim intColWidth As Integer
Dim strColText As String
    'Error Trap
    On Error GoTo AutoSizeColumn_ErrorHandler
    'For the given column go down it, find the widest bit of data and size the column
    intColWidth = 0
    For intRowCounter = 0 To grdControl.Rows - 1
        strColText = grdControl.TextMatrix(intRowCounter, intColumn)
        If grdControl.Parent.TextWidth(strColText) > intColWidth Then
            intColWidth = grdControl.Parent.TextWidth(strColText)
        End If
    Next
    grdControl.ColWidth(intColumn) = intColWidth + 120
Exit Sub
'Error Handler
AutoSizeColumn_ErrorHandler:
    Exit Sub
End Sub

Public Function WriteGridToCSV(ByVal strFileName As String, ByRef grdControl As MSFlexGrid, ByVal booWriteColumnHeaders As Boolean, ByVal booWriteRowHeaders As Boolean) As Boolean
'Created by Robin G Brown, 6/5/97
'Write the contents of the control to a CSV file, _
    assumes : strfilename is valid, contains a path and does not exist!
'Function Declares
Const conFunction = "WriteGridToCSV"
Dim intFileNumber As Integer
Dim intColStart As Integer
Dim strCSV As String
    'Error Trap
    On Error GoTo WriteGridToCSV_ErrorHandler
    'Open the file
    intFileNumber = FreeFile
    Open strFileName For Output As #intFileNumber
    'Check to see if we need row headers
    If booWriteRowHeaders = True Then
        intColStart = 0
    Else
        intColStart = 1
    End If
    'Write column headers
    If booWriteColumnHeaders = True Then
        strCSV = ""
        For intColCounter = intColStart To grdControl.Cols - 1
            strCSV = strCSV & grdControl.TextMatrix(0, intColCounter) & ", "
        Next
        'Trim the last ", "
        strCSV = Left$(strCSV, Len(strCSV) - 2)
        'Write the line
        Print #intFileNumber, strCSV & vbCrLf
    End If
    'Write data rows
    For intRowCounter = 1 To grdControl.Rows - 1
        strCSV = ""
        For intColCounter = intColStart To grdControl.Cols - 1
            'Write the data
            strCSV = strCSV & grdControl.TextMatrix(intRowCounter, intColCounter) & ", "
        Next
        'Trim the last ", "
        strCSV = Left$(strCSV, Len(strCSV) - 2)
        'Write the line
        Print #intFileNumber, strCSV & vbCrLf
    Next
'Close the file
    Close #intFileNumber
    WriteGridToCSV = True
Exit Function
'Error Handler
WriteGridToCSV_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        'MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conFunction
    End Select
    On Error Resume Next
    'Close the file
    Close #intFileNumber
    WriteGridToCSV = False
    Exit Function
End Function

Public Sub FillArrayFromGrid(ByRef grdControl As MSFlexGrid, ByRef varArray As Variant, ByVal booIncludeHeaders As Boolean)
'Created by Robin G Brown, 6/5/97
'Fill an array with data from the grid
'Sub Declares
Const conSub = "FillArrayFromGrid"
Dim intLBound As Integer
'Error Trap
    On Error GoTo FillArrayFromGrid_ErrorHandler
    If booIncludeHeaders = True Then
        intLBound = 0
    Else
        intLBound = 1
    End If
    ReDim varArray(intLBound To grdControl.Rows - 1, intLBound To grdControl.Cols - 1)
    'Read column headers
    If booIncludeHeaders = True Then
        For intColCounter = intLBound To grdControl.Cols - 1
            varArray(0, intColCounter) = grdControl.TextMatrix(0, intColCounter)
        Next
    End If
    'Read data rows
    For intRowCounter = 1 To grdControl.Rows - 1
        For intColCounter = intLBound To grdControl.Cols - 1
            'Read the data
            varArray(intRowCounter, intColCounter) = grdControl.TextMatrix(intRowCounter, intColCounter)
        Next
    Next
Exit Sub
'Error Handler
FillArrayFromGrid_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        'MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Public Sub FillGridFromArray(ByRef grdControl As MSFlexGrid, ByRef varArray() As Variant)
'Created by Robin G Brown, 2/5/97
'Fill a grid from an array
'Sub Declares
Const conSub = "FillGridFromArray"
Dim intColCounter As Integer
Dim intRowCounter As Integer
    'Error Trap
    On Error GoTo FillGridFromArray_ErrorHandler
    'Set the number of cols and rows
    grdControl.FixedCols = 0
    grdControl.FixedRows = 0
    grdControl.Cols = UBound(varArray, 2)
    grdControl.Rows = UBound(varArray, 1)
    'Write the col headers
    For intColCounter = 1 To grdControl.Cols - 1
        grdControl.TextMatrix(0, intColCounter) = "Col " & intColCounter
    Next
    For intRowCounter = 1 To grdControl.Rows - 1
        'Write the row header
        grdControl.TextMatrix(intRowCounter, 0) = "Row " & intRowCounter
        For intColCounter = 1 To grdControl.Cols - 1
            'Write the data
            grdControl.TextMatrix(intRowCounter, intColCounter) = _
                varArray(intRowCounter, intColCounter)
        Next
    Next
    'Set the number of fixed cols and rows
    grdControl.FixedCols = 1
    grdControl.FixedRows = 1
Exit Sub
'Error Handler
FillGridFromArray_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Public Sub InsertRow(ByRef grdControl As MSFlexGrid, ByVal intRow As Integer)
'Created by Robin G Brown, 2/5/97
'Inserts a row above the given row in the given control, _
    assumes that there are either no fixed rows/cols, _
    or that they will be corrected in another way
'Function Declares
Const conFunction = "InsertRow"
Dim strText As String
    'Error Trap
    On Error GoTo InsertRow_ErrorHandler
    grdControl.Rows = grdControl.Rows + 1
    'Working from the 2nd bottom row upwards, copy each row into the one below it
    For intRowCounter = grdControl.Rows - 2 To intRow Step -1
        For intColCounter = 0 To grdControl.Cols - 1
            grdControl.Col = intColCounter
            grdControl.Row = intRowCounter
            strText = grdControl.Text
            grdControl.Row = intRowCounter + 1
            grdControl.Text = strText
        Next
    Next
    'Blank out the 'inserted' row
    grdControl.Row = intRow
    For intColCounter = 0 To grdControl.Cols - 1
        grdControl.Col = intColCounter
        grdControl.Text = ""
    Next
Exit Sub
'Error Handler
InsertRow_ErrorHandler:
    'MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conFunction
    Exit Sub
End Sub

Public Sub InsertCol(ByRef grdControl As MSFlexGrid, ByVal intColumn As Integer)
'Created by Robin G Brown, 2/5/97
'Inserts a column before the given column in the given control, _
    assumes that there are either no fixed rows/cols, _
    or that they will be corrected in another way
'Function Declares
Const conFunction = "InsertCol"
Dim strText As String
    'Error Trap
    On Error GoTo InsertCol_ErrorHandler
    grdControl.Cols = grdControl.Cols + 1
    'Working from the 2nd last column leftwards, copy each col into the one to the right of it
    For intColCounter = grdControl.Cols - 2 To intColumn Step -1
        For intRowCounter = 0 To grdControl.Rows - 1
            grdControl.Col = intColCounter
            grdControl.Row = intRowCounter
            strText = grdControl.Text
            grdControl.Col = intColCounter + 1
            grdControl.Text = strText
        Next
    Next
    'Blank out the 'inserted' column
    grdControl.Col = intColumn
    For intRowCounter = 0 To grdControl.Rows - 1
        grdControl.Row = intRowCounter
        grdControl.Text = ""
    Next
Exit Sub
'Error Handler
InsertCol_ErrorHandler:
    'MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conFunction
    Exit Sub
End Sub

Public Sub SelectItems(ByRef grdControl As MSFlexGrid, ByRef fgcCell As FlexGridCell, ByVal intInstruction As Integer)
'Created by Robin G Brown, 13/5/97
'Select various elements of the grid by colouring them differently
'Sub Declares
Const conSub = "SelectItems"
Dim intOldFillStyle As Integer
    'Error Trap
    On Error GoTo SelectItems_ErrorHandler
    'Start by clearing any previous Selection
    intOldFillStyle = grdControl.FillStyle
    grdControl.FillStyle = flexFillRepeat
    grdControl.Row = grdControl.FixedRows
    grdControl.Col = grdControl.FixedCols
    grdControl.RowSel = grdControl.Rows - 1
    grdControl.ColSel = grdControl.Cols - 1
    grdControl.CellForeColor = grdControl.ForeColor
    grdControl.CellBackColor = grdControl.BackColor
    'Then colour in cells according to the fgcCell and the SelectItemsion instruction
    Select Case intInstruction
    Case conSelectClear
        'Do nothing
    Case conSelectColumn
        grdControl.Col = fgcCell.GridCol
        grdControl.Row = grdControl.FixedRows
        grdControl.ColSel = fgcCell.GridCol
        grdControl.RowSel = grdControl.Rows - 1
        grdControl.CellForeColor = grdControl.ForeColorSel
        grdControl.CellBackColor = grdControl.BackColorSel
    Case conSelectRow
        grdControl.Col = grdControl.FixedCols
        grdControl.Row = fgcCell.GridRow
        grdControl.ColSel = grdControl.Cols - 1
        grdControl.RowSel = fgcCell.GridRow
        grdControl.CellForeColor = grdControl.ForeColorSel
        grdControl.CellBackColor = grdControl.BackColorSel
    Case conSelectCell
        grdControl.FillStyle = flexFillSingle
        grdControl.Col = grdControl.FixedCols
        grdControl.Row = fgcCell.GridRow
        grdControl.CellForeColor = grdControl.ForeColorSel
        grdControl.CellBackColor = grdControl.BackColorSel
    Case Else
        'Do nothing
    End Select
    grdControl.FillStyle = intOldFillStyle
Exit Sub
'Error Handler
SelectItems_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Public Sub GenericFillGrid(ByRef grdControl As MSFlexGrid)
'Created by Robin G Brown, 2/5/97
'Default behaviour
'Sub Declares
Const conSub = "GenericFillGrid"
    'Error Trap
    On Error GoTo GenericFillGrid_ErrorHandler
    'Set the number of cols and rows
    grdControl.Cols = 1
    grdControl.Rows = 1
    'Write the col headers
    grdControl.Row = 0
    For intColCounter = 1 To grdControl.Cols - 1
        grdControl.TextMatrix(0, intColCounter) = "Column Header"
    Next
    For intRowCounter = 1 To grdControl.Rows - 1
        'Write the row header
        grdControl.TextMatrix(intRowCounter, 0) = "Row Header"
        For intColCounter = 1 To grdControl.Cols - 1
            'Write the data
            grdControl.TextMatrix(intRowCounter, intColCounter) = "Data"
        Next
    Next
    'Set the number of fixed cols and rows
    grdControl.FixedCols = 1
    grdControl.FixedRows = 1
Exit Sub
'Error Handler
GenericFillGrid_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub


