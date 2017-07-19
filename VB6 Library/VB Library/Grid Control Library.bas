Attribute VB_Name = "modGridLib"
'Created by Robin G Brown, 18/4/97
'-----------------------------------------------------------------------------
'   This module contains code for _
    handling grid controls
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "modGridLib"
Private intColCounter As Integer
Private intRowCounter As Integer
Private intReturn As Integer
'This type can be used to return _
    multiple data values from a function that interrogates a grid
'Public Type GridCell
'    'Use gdc as a prefix
'    GridCol As Integer
'    GridRow As Integer
'    GridText As String
'End Type
'-----------------------------------------------------------------------------
'   Spoiler information for Grids
'
'   1   2   3   4   5   6   7   8   9
'   1   2   3   4   5   6   7   8   9
'   1   2   3   4   5   6   7   8   9
'   1   2   3   4   5   6   7   8   9
'
'-----------------------------------------------------------------------------

Public Sub FillArrayFromGrid(ByRef grdControl As Grid, ByRef varArray As Variant, ByVal booIncludeHeaders As Boolean)
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
    ReDim varArray(intLBound To grdControl.Cols - 1, intLBound To grdControl.Rows - 1)
    'Read column headers
    If booIncludeHeaders = True Then
        grdControl.Row = 0
        For intColCounter = intLBound To grdControl.Cols - 1
            grdControl.Col = intColCounter
            varArray(0, intColCounter) = grdControl.Text
        Next
    End If
    'Read data rows
    For intRowCounter = 1 To grdControl.Rows - 1
        grdControl.Row = intRowCounter
        For intColCounter = intLBound To grdControl.Cols - 1
            grdControl.Col = intColCounter
            'Read the data
            varArray(intRowCounter, intColCounter) = grdControl.TextMatrix
        Next
    Next
Exit Sub
'Error Handler
FillArrayFromGrid_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Public Function WriteGridToCSV(ByVal strFileName As String, ByRef grdControl As Grid, ByVal booWriteColumnHeaders As Boolean, ByVal booWriteRowHeaders As Boolean) As Boolean
'Created by Robin G Brown, 6/5/97
'Write the contents of the control to a CSV file, _
    assumes : strfilename is valid, contains a path and does not exist!
'Function Declares
Const conFunction = "WriteGridToCSV"
Dim intFilenumber As Integer
Dim intColStart As Integer
Dim strCSV As String
    'Error Trap
    On Error GoTo WriteGridToCSV_ErrorHandler
    'Open the file
    intFilenumber = FreeFile
    Open strFileName For Output As #intFilenumber
    'Check to see if we need row headers
    If booWriteRowHeaders = True Then
        intColStart = 0
    Else
        intColStart = 1
    End If
    'Write column headers
    If booWriteColumnHeaders = True Then
        strCSV = ""
        grdControl.Row = 0
        For intColCounter = intColStart To grdControl.Cols - 1
            grdControl.Col = intColCounter
            strCSV = strCSV & grdControl.Text & ", "
        Next
        'Trim the last ", "
        strCSV = Left$(strCSV, Len(strCSV) - 2)
        'Write the line
        Print #intFilenumber, strCSV & vbCrLf
    End If
    'Write data rows
    For intRowCounter = 1 To grdControl.Rows - 1
        strCSV = ""
        grdControl.Row = intRowCounter
        For intColCounter = intColStart To grdControl.Cols - 1
            'Write headers, if required
            grdControl.Col = intColCounter
            If intColCounter = 0 Then
                strCSV = strCSV & grdControl.Text & ", "
            Else
                'Then write the data
                strCSV = strCSV & grdControl.Text & ", "
            End If
        Next
        'Trim the last ", "
        strCSV = Left$(strCSV, Len(strCSV) - 2)
        'Write the line
        Print #intFilenumber, strCSV & vbCrLf
    Next
'Close the file
    Close #intFilenumber
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
    Close #intFilenumber
    WriteGridToCSV = False
    Exit Function
End Function

Public Sub GenericFillGrid(ByRef grdControl As Grid)
'Created by Robin G Brown, 2/5/97
'Default behaviour
'Sub Declares
Const conSub = "GenericFillGrid"
Dim intColCounter As Integer
Dim intRowCounter As Integer
    'Error Trap
    On Error GoTo GenericFillGrid_ErrorHandler
    'Set the number of cols and rows
    grdControl.Cols = 1
    grdControl.Rows = 1
    'Write the col headers
    grdControl.Row = 0
    For intColCounter = 1 To grdControl.Cols - 1
        grdControl.Col = intColCounter
        grdControl.Text = "Column Header"
    Next
    For intRowCounter = 1 To grdControl.Rows - 1
        'Write the row header
        grdControl.Col = 0
        grdControl.Text = "Row Header"
        For intColCounter = 1 To grdControl.Cols - 1
            'Write the data
            grdControl.Col = intColCounter
            grdControl.Text = "Data"
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

Public Sub InsertRow(ByRef grdControl As Grid, ByVal intRow As Integer)
'Created by Robin G Brown, 18/4/97
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

Public Sub InsertCol(ByRef grdControl As Grid, ByVal intColumn As Integer)
'Created by Robin G Brown, 18/4/97
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

Public Function FillGridFromRecordset(dynaSource As Recordset, ctlGrid As Control) As Integer
'RGB/17/1/96
'This function fills a grid with data from a recordset supplied to it
    
Dim intFields As Integer
Dim intFieldCounter As Integer
Dim intRowCounter As Integer
Dim intColWidth() As Integer

    On Error GoTo FillGridFromRecordset_ErrorHandler
    FillGridFromRecordset = False
    'Check that target is a grid and that there are records to fill it with
    If Not TypeOf ctlGrid Is Grid Then
        'not a grid so ignore it
        FillGridFromRecordset = False
        Exit Function
    End If
    dynaSource.MoveLast
    If dynaSource.RecordCount = 0 Or dynaSource.RecordCount > 16000 Then
        FillGridFromRecordset = False
        Exit Function
    End If
    'Everything OK so set some variables
    dynaSource.MoveFirst
    intFields = dynaSource.Fields.Count
    intRowCounter = 0
    ReDim intColWidth(1 To intFields + 1)
    'Create Grid rows/cols
    ctlGrid.FixedRows = 0
    ctlGrid.FixedCols = 0
    ctlGrid.Rows = 1
    ctlGrid.Cols = intFields + 1
    'Set top left corner cell
    ctlGrid.Row = 0
    ctlGrid.Col = 0
    ctlGrid.Text = dynaSource.Name
    'Fill Grid Headings
    For intFieldCounter = 1 To intFields
        ctlGrid.Col = intFieldCounter
        ctlGrid.Text = CStr(dynaSource.Fields(intFieldCounter - 1).Name)
        intColWidth(intFieldCounter) = Len(CStr(dynaSource.Fields(intFieldCounter - 1).Name))
    Next
    intRowCounter = intRowCounter + 1
    'Fill data
    Do While Not dynaSource.EOF
        ctlGrid.Rows = ctlGrid.Rows + 1
        ctlGrid.Row = ctlGrid.Row + 1
        ctlGrid.Col = 0
        ctlGrid.Text = CStr(intRowCounter)
        For intFieldCounter = 1 To intFields
            ctlGrid.Col = intFieldCounter
            ctlGrid.Text = CStr(dynaSource.Fields(intFieldCounter - 1))
        Next
        intRowCounter = intRowCounter + 1
        dynaSource.MoveNext
    Loop
    'Set height/width of cols (in twips)
    If intRowCounter < 1000 Then
        ctlGrid.ColWidth(0) = ctlGrid.Parent.TextWidth("0000")
    Else
        ctlGrid.ColWidth(0) = ctlGrid.Parent.TextWidth("00000")
    End If
    For intFieldCounter = 1 To intFields
        ctlGrid.Col = intFieldCounter
        ctlGrid.ColWidth(intFieldCounter) = ctlGrid.Parent.TextWidth("A") * intColWidth(intFieldCounter)
    Next
    'Set fixed rows/cols and any other properties required
    ctlGrid.FixedRows = 1
    ctlGrid.FixedCols = 1
    FillGridFromRecordset = intRowCounter
    
Exit Function
FillGridFromRecordset_ErrorHandler:
    FillGridFromRecordset = False
    Exit Function
    
End Function


