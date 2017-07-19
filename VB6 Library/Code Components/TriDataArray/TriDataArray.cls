VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TriData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Created by Robin G Brown, 6th May 1997
'-----------------------------------------------------------------------------
'   This class module contains code for _
    holding and manipulating the data for the grid
'   Use dta as a prefix for this class
'-----------------------------------------------------------------------------
'Spoiler information for TriDataArray
'
'   1. The various arrays in TriDataArray represent a grid of data with header and formatting information
'   2. The index is always (row, column) for all arrays
'   3. curData() contains currency data figures
'   4. curData() starts from (1,1) at all times
'   5. strColumnHeaders() contains column header data in two dimensions
'   6. strColumnHeaders() starts from (0,1) at all times
'   7. Row(0) in strColumnHeaders() contains UWYear headers
'   8. Row(1) in strColumnHeaders() contains Account headers
'   9. Columns in strColumnHeaders() are grouped into years with
'      a size equal to conAccounts
'   10. Total number of data columns is intUWYears * conAccounts (numbered from 1)
'   11. accAccounts() is used to determine
'       whether various display elements should be displayed FOR EACH UWYEAR
'   12. strRowHeaders() contains row header data in one dimension
'   13. strRowHeaders() starts from (1,0) at all times
'   14. Col(0) in strRowHeaders() contains Year headers
'   15A. Col(1) in strRowHeaders() contains Quarter headers
'   15B. Col(2) in strRowHeaders() contains Month headers
'   16. Rows in strRowHeaders() are grouped into years with
'       a size equal to 12
'   17. Total number of data rows is equal to intYears * conMonths (numbered from 1)
'   19. Accounts/Columns/Elements are hidden by setting the width of the control's column
'       for that account/column/element to 0
'   20. vfaFlexArray contains a reference to a VSFlexArray control that will display the data
'   21. graGraphControl contains a reference to an MSChart control that will graph the data
'   22. Graphing can be done for single columns or for all columns that represent an account
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
'Private types
Private Type Account
    'Use acc as a prefix
    Name As String
    Display As Boolean
    Mask As String
End Type
Private Type DataSource
    'Use dts as a prefix
    SourceType As String
    ConnectString As String
    SQL As String
End Type
Private Type GridMousePosition
    'Use grm as a prefix
    Col As Long
    Row As Long
End Type
'Private properties
Private dtsDataSource As DataSource
Private strColumnHeaders() As String
Private strRowHeaders() As String
Private curData() As Currency
Private accAccounts() As Account
Private intUWYears As Integer
Private intYears As Integer
Private intReturn As Integer
Private strInstance As String
Private intSelectedColumn As Integer
Private grmGridMouse As GridMousePosition
'Private bytSubTotals As Byte
'Hold references to specific controls for displaying and graphing data
Private vfaFlexArray As vsFlexArray
Private graGraphControl As MSChart
'Counter variables
Private lngRowCounter As Long
Private lngColCounter As Long
Private intAcCounter As Integer
Private intMonthcounter As Integer
'Constants
Private Const conClass = "TriDataArray.TriData"
Private Const conAccounts = 9
Private Const conQuarters = 4
Private Const conMonths = 12
Private Const conNumFixedCols = 3
Private Const conNumFixedRows = 2
'File constants
Private Const conSeperator = ","
Private Const conColon = ":"
Private Const conUnknownLine = -1
Private Const conColHead1Line = 1
Private Const conColHead2Line = 2
Private Const conRowHead1Line = 3
Private Const conRowHead2Line = 4
Private Const conAccNameLine = 5
Private Const conAccDisplayLine = 6
Private Const conAccColourLine = 7
Private Const conAccFormatLine = 8
Private Const conDataLine = 9
Private Const conColHeaderLine = 10
Private Const conRowHeaderLine = 11
Private Const conFormatHeaderLine = 12
Private Const conDataHeaderLine = 13
'Error Constants
Private Const conBadYearValueError = vbObjectError + 1
Private Const conCannotHideLastAccountError = vbObjectError + 2
Private Const conInvalidObjectError = vbObjectError + 3

Private Sub Class_Initialize()
'Created by Robin G Brown, 6/5/97
'Default Behaviour
'Sub Declares
    'Error Trap
    On Error Resume Next
    'Initialise the arrays and default values
    strInstance = "Undefined"
    intUWYears = 15
    intYears = 15
    intSelectedColumn = -1
    Set vfaFlexArray = Nothing
    Set graGraphControl = Nothing
    With dtsDataSource
        .SourceType = "NO_SOURCE_TYPE_FOR_THIS_DATA"
        .ConnectString = "NO_CONNECT_STRING_FOR_THIS_DATA"
        .SQL = "NO_SQL_FOR_THIS_DATA"
    End With
    Call InitialiseData
    Call InitialiseFormat
End Sub

'-----------------------------------------------------------------------------
'   Public Functions and Methods
'-----------------------------------------------------------------------------

Public Sub FileSave(ByRef sffSectionFile As SecFile, ByVal strSaveFileName As String)
'Created by Robin G Brown, 19/5/97
'Save the data and formatting to a file
'Sub Declares
Const conSub = "FileSave"
Dim intFileNumber As Integer
Dim intDataCol As Integer
Dim intDataRow As Integer
Dim strDataLine As String
Dim strSecondDataLine As String
Dim strThirdDataLine As String
Dim lngYearcounter As Long
    'Error Trap
    On Error GoTo FileSave_ErrorHandler
    'Save the header data
    With sffSectionFile
        'Save window data
        .AddSection "WINDOW"
        .SetSectionSetting "WINDOW", "state", vfaFlexArray.Parent.WindowState
        .SetSectionSetting "WINDOW", "left", vfaFlexArray.Parent.Left
        .SetSectionSetting "WINDOW", "top", vfaFlexArray.Parent.Top
        .SetSectionSetting "WINDOW", "height", vfaFlexArray.Parent.Height
        .SetSectionSetting "WINDOW", "width", vfaFlexArray.Parent.Width
        'Save the formatting data
        .AddSection "DATASOURCE"
        .SetSectionSetting "DATASOURCE", "sourcetype", dtsDataSource.SourceType
        .SetSectionSetting "DATASOURCE", "connectstring", dtsDataSource.ConnectString
        .SetSectionSetting "DATASOURCE", "sql", dtsDataSource.SQL
        'Column headers
        .AddSection "COLUMNHEADER"
        .SetSectionSetting "COLUMNHEADER", "fixedcols", conNumFixedCols
        .SetSectionSetting "COLUMNHEADER", "numuwyears", intUWYears
        strDataLine = ""
        strSecondDataLine = ""
        For lngColCounter = 1 To intUWYears
            For intAcCounter = 1 To conAccounts
                intDataCol = ((lngColCounter - 1) * conAccounts) + intAcCounter
                strDataLine = strDataLine & strColumnHeaders(0, intDataCol) & conSeperator
                strSecondDataLine = strSecondDataLine & strColumnHeaders(1, intDataCol) & conSeperator
            Next
        Next
        strDataLine = Left$(strDataLine, Len(strDataLine) - 1) & ""
        strSecondDataLine = Left$(strSecondDataLine, Len(strSecondDataLine) - 1) & ""
        .SetSectionSetting "COLUMNHEADER", "uwyears", strDataLine
        .SetSectionSetting "COLUMNHEADER", "accounts", strSecondDataLine
        'Row headers
        .AddSection "ROWHEADER"
        .SetSectionSetting "ROWHEADER", "fixedrows", conNumFixedRows
        .SetSectionSetting "ROWHEADER", "numyears", intYears
        strDataLine = ""
        strSecondDataLine = ""
        strThirdDataLine = ""
        For lngRowCounter = 1 To intYears
            For intMonthcounter = 1 To conMonths
                intDataRow = ((lngRowCounter - 1) * conMonths) + intMonthcounter
                strDataLine = strDataLine & strRowHeaders(intDataRow, 0) & conSeperator
                strSecondDataLine = strSecondDataLine & strRowHeaders(intDataRow, 1) & conSeperator
                strThirdDataLine = strThirdDataLine & strRowHeaders(intDataRow, 2) & conSeperator
            Next
        Next
        strDataLine = Left$(strDataLine, Len(strDataLine) - 1) & ""
        strSecondDataLine = Left$(strSecondDataLine, Len(strSecondDataLine) - 1) & ""
        strThirdDataLine = Left$(strThirdDataLine, Len(strThirdDataLine) - 1) & ""
        .SetSectionSetting "ROWHEADER", "years", strDataLine
        .SetSectionSetting "ROWHEADER", "quarters", strSecondDataLine
        .SetSectionSetting "ROWHEADER", "months", strThirdDataLine
        'Save the formatting data
        .AddSection "DATAFORMAT"
        'Account names
        strDataLine = ""
        For lngColCounter = 1 To UBound(accAccounts, 1)
            strDataLine = strDataLine & accAccounts(lngColCounter).Name & conSeperator
        Next
        strDataLine = Left$(strDataLine, Len(strDataLine) - 1) & ""
        .SetSectionSetting "DATAFORMAT", "names", strDataLine
        'Display accounts
        strDataLine = ""
        For lngColCounter = 1 To UBound(accAccounts, 1)
            strDataLine = strDataLine & accAccounts(lngColCounter).Display & conSeperator
        Next
        strDataLine = Left$(strDataLine, Len(strDataLine) - 1) & ""
        .SetSectionSetting "DATAFORMAT", "display", strDataLine
        'Account format masks
        strDataLine = ""
        For lngColCounter = 1 To UBound(accAccounts, 1)
            strDataLine = strDataLine & accAccounts(lngColCounter).Mask & conSeperator
        Next
        strDataLine = Left$(strDataLine, Len(strDataLine) - 1) & ""
        .SetSectionSetting "DATAFORMAT", "masks", strDataLine
        'Save the actual data
        .AddSection "DATA"
        .SetSectionSetting "DATA", "numcols", intUWYears * UBound(accAccounts, 1)
        .SetSectionSetting "DATA", "numrows", intYears * conMonths
        For lngYearcounter = 1 To intYears
            For lngRowCounter = 1 To conMonths
                strDataLine = ""
                For lngColCounter = 1 To intUWYears
                    For intAcCounter = 1 To UBound(accAccounts, 1)
                        intDataCol = ((lngColCounter - 1) * conAccounts) + intAcCounter
                        intDataRow = ((lngYearcounter - 1) * conMonths) + lngRowCounter
                        strDataLine = strDataLine & curData(intDataRow, intDataCol) & conSeperator
                    Next
                Next
                strDataLine = Left$(strDataLine, Len(strDataLine) - 1) & ""
                .SetSectionSetting "DATA", "row" & intDataRow, strDataLine
            Next
        Next
        .FileSave strSaveFileName
    End With
Exit Sub
'Error Handler
FileSave_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conClass & ":" & conSub, "Unexpected error : " & Err.Description
    End Select
    Exit Sub
End Sub

Public Sub FileOpen(ByRef sffSectionFile As SecFile, ByVal strOpenFileName As String)
'Created by Robin G Brown, 19/5/97
'Retrieve the formatting and data from a file
'Sub Declares
Const conSub = "FileOpen"
Dim intFileNumber As Integer
Dim intLineType As Integer
Dim strDataLine As String
Dim intDataLineCounter As Integer
Dim strLineArray() As String
Dim lngSectionCounter As Long
Dim lngNumDataCols As Long
Dim lngNumDataRows As Long
    'Error Trap
    On Error GoTo FileOpen_ErrorHandler
    With sffSectionFile
        'Open the file
        .FileOpen strOpenFileName
        Me.Name = strOpenFileName
        For lngSectionCounter = 1 To .Sections
            Select Case .IndexSection(lngSectionCounter)
            Case "DATA"
                'Read data into curdata()
                lngNumDataCols = .GetSectionSetting("DATA", "numcols")
                lngNumDataRows = .GetSectionSetting("DATA", "numrows")
                ReDim curData(1 To lngNumDataRows, 1 To lngNumDataCols)
                ReDim strLineArray(1 To lngNumDataCols)
                For lngRowCounter = 1 To lngNumDataRows
                    strDataLine = .GetSectionSetting("DATA", "row" & lngRowCounter)
                    'Read CSV line into array
                    intReturn = WriteCSVToArray(strDataLine, strLineArray())
                    For lngColCounter = 1 To lngNumDataCols
                        curData(lngRowCounter, lngColCounter) = CCur(strLineArray(lngColCounter))
                    Next
                Next
            Case "COLUMNHEADER"
                'Retrieve the UWYear header data
                intUWYears = CInt(.GetSectionSetting("COLUMNHEADER", "numuwyears"))
                ReDim strColumnHeaders(0 To 1, 1 To intYears * conAccounts)
                strDataLine = .GetSectionSetting("COLUMNHEADER", "uwyears")
                ReDim strLineArray(1 To CharactersInString(strDataLine, conSeperator) + 1)
                intReturn = WriteCSVToArray(strDataLine, strLineArray())
                For lngColCounter = 1 To UBound(strLineArray, 1)
                    strColumnHeaders(0, lngColCounter) = strLineArray(lngColCounter)
                Next
                'Retrieve the Account header data
                strDataLine = .GetSectionSetting("COLUMNHEADER", "accounts")
                ReDim strLineArray(1 To CharactersInString(strDataLine, conSeperator) + 1)
                intReturn = WriteCSVToArray(strDataLine, strLineArray())
                For lngColCounter = 1 To UBound(strLineArray, 1)
                    strColumnHeaders(1, lngColCounter) = strLineArray(lngColCounter)
                Next
            Case "ROWHEADER"
                'Retrieve the year header data
                intYears = CInt(.GetSectionSetting("ROWHEADER", "numyears"))
                ReDim strRowHeaders(1 To intYears * conMonths, 0 To 2)
                strDataLine = .GetSectionSetting("ROWHEADER", "years")
                ReDim strLineArray(1 To CharactersInString(strDataLine, conSeperator) + 1)
                intReturn = WriteCSVToArray(strDataLine, strLineArray())
                For lngRowCounter = 1 To UBound(strLineArray, 1)
                    strRowHeaders(lngRowCounter, 0) = strLineArray(lngRowCounter)
                Next
                'Retrieve the quarter header data
                strDataLine = .GetSectionSetting("ROWHEADER", "quarters")
                ReDim strLineArray(1 To CharactersInString(strDataLine, conSeperator) + 1)
                intReturn = WriteCSVToArray(strDataLine, strLineArray())
                For lngRowCounter = 1 To UBound(strLineArray, 1)
                    strRowHeaders(lngRowCounter, 1) = strLineArray(lngRowCounter)
                Next
                'Retrieve the month header data
                strDataLine = .GetSectionSetting("ROWHEADER", "months")
                ReDim strLineArray(1 To CharactersInString(strDataLine, conSeperator) + 1)
                intReturn = WriteCSVToArray(strDataLine, strLineArray())
                For lngRowCounter = 1 To UBound(strLineArray, 1)
                    strRowHeaders(lngRowCounter, 2) = strLineArray(lngRowCounter)
                Next
            Case "DATAFORMAT"
                'Retrieve the name formatting data
                strDataLine = .GetSectionSetting("DATAFORMAT", "names")
                ReDim strLineArray(1 To CharactersInString(strDataLine, conSeperator) + 1)
                intReturn = WriteCSVToArray(strDataLine, strLineArray())
                For lngColCounter = 1 To UBound(strLineArray, 1)
                    accAccounts(lngColCounter).Name = strLineArray(lngColCounter)
                Next
                'Retrieve the display formatting data
                strDataLine = .GetSectionSetting("DATAFORMAT", "display")
                ReDim strLineArray(1 To CharactersInString(strDataLine, conSeperator) + 1)
                intReturn = WriteCSVToArray(strDataLine, strLineArray())
                For lngColCounter = 1 To UBound(strLineArray, 1)
                    accAccounts(lngColCounter).Display = strLineArray(lngColCounter)
                Next
                'Retrieve the mask formatting data
                strDataLine = .GetSectionSetting("DATAFORMAT", "masks")
                ReDim strLineArray(1 To CharactersInString(strDataLine, conSeperator) + 1)
                intReturn = WriteCSVToArray(strDataLine, strLineArray())
                For lngColCounter = 1 To UBound(strLineArray, 1)
                    accAccounts(lngColCounter).Mask = strLineArray(lngColCounter)
                Next
            Case "DATASOURCE"
                dtsDataSource.SourceType = .GetSectionSetting("DATASOURCE", "sourcetype")
                dtsDataSource.ConnectString = .GetSectionSetting("DATASOURCE", "connectstring")
                dtsDataSource.SQL = .GetSectionSetting("DATASOURCE", "sql")
            Case "WINDOW"
                vfaFlexArray.Parent.Left = .GetSectionSetting("WINDOW", "left")
                vfaFlexArray.Parent.Top = .GetSectionSetting("WINDOW", "top")
                vfaFlexArray.Parent.Height = .GetSectionSetting("WINDOW", "height")
                vfaFlexArray.Parent.Width = .GetSectionSetting("WINDOW", "width")
                vfaFlexArray.Parent.WindowState = .GetSectionSetting("WINDOW", "state")
            Case Else
                'Ignore this section
            End Select
        Next
    End With
Exit Sub
'Error Handler
FileOpen_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conClass & ":" & conSub, "Unexpected error : " & Err.Description
    End Select
    Close #intFileNumber
    Exit Sub
End Sub

Public Function ShowAccount(ByVal intAccount As Integer, booVisible As Boolean)
'Created by Robin G Brown, 7/5/97
'Sets accAccounts(intAccount) according to booVisible
'Sub Declares
Dim intVisibleAccounts As Integer
    'Error Trap
    On Error Resume Next
    If booVisible = True Then
        accAccounts(intAccount).Display = True
    Else
        'Check that the last visible account is not being hidden
        intVisibleAccounts = 0
        For intAcCounter = 1 To conAccounts
            If accAccounts(intAcCounter).Display = True Then
                intVisibleAccounts = intVisibleAccounts + 1
            End If
        Next
        If intVisibleAccounts > 1 Then
            accAccounts(intAccount).Display = False
        Else
            Err.Raise conCannotHideLastAccountError, conClass, "Cannot hide all accounts"
        End If
    End If
End Function

Public Function GetAccountFromColumn(ByVal intColumn As Integer) As Integer
'Created by Robin G Brown, 13/5/97
'Return which account is in a given grid column
'Function Declares
    'Error Trap
    On Error Resume Next
    intColumn = intColumn - (conNumFixedCols - 1)
    GetAccountFromColumn = (intColumn Mod conAccounts)
End Function

Public Sub GraphOnAccount(ByVal intGridColumn As Integer)
'Created by Robin G Brown, 6/5/97
'Set up a chart from the curData() array, graphing all columns for the given account
'Sub Declares
Const conSub = "GraphOnAccount"
Dim intSelectedAccount As Integer
Dim intNewColumn As Integer
    'Error Trap
    On Error GoTo GraphOnAccount_ErrorHandler
    'Set up the chart control before adding data
    Call PrepareChart
    intSelectedAccount = GetAccountFromColumn(intGridColumn)
    For lngColCounter = 1 To intUWYears
        intNewColumn = ((lngColCounter - 1) * conAccounts) + intSelectedAccount + conNumFixedCols - 1
        Call AddColumnToGraph(intNewColumn)
    Next
Exit Sub
'Error Handler
GraphOnAccount_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conClass & ":" & conSub, "Unexpected error : " & Err.Description
    End Select
    Exit Sub
End Sub

Public Sub GraphOnColumn(ByVal intGridColumn As Integer)
'Created by Robin G Brown, 6/5/97
'Set up a chart from the curData() array, _
    graphing only the column indicated by intGridColumn
'Sub Declares
Const conSub = "GraphOnColumn"
    'Error Trap
    On Error GoTo GraphOnColumn_ErrorHandler
    Call PrepareChart
    Call AddColumnToGraph(intGridColumn)
Exit Sub
'Error Handler
GraphOnColumn_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conClass & ":" & conSub, "Unexpected error : " & Err.Description
    End Select
    Exit Sub
End Sub

Public Sub AddMultipleColumnsToGraph(ByVal intGridColumn As Integer)
'Created by Robin G Brown, 20/8/97
'Insert a number of columns of data into the graph, utilising the AddColumnToGraph method
'Function Declares
Const conSub = "AddMultipleColumnsToGraph"
Dim lngUWYearCounter As Long
Dim intLastColumn As Integer
Dim intSelectedAccount As Integer
Dim intNewColumn As Integer
    'Error Trap
    On Error GoTo AddMultipleColumnsToGraph_ErrorHandler
    intSelectedAccount = GetAccountFromColumn(intGridColumn)
    For lngUWYearCounter = 1 To intUWYears
        intNewColumn = ((lngUWYearCounter - 1) * conAccounts) + intSelectedAccount + conNumFixedCols - 1
        Call AddColumnToGraph(intNewColumn)
    Next
Exit Sub
'Error Handler
AddMultipleColumnsToGraph_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conClass & ":" & conSub, "Unexpected error : " & Err.Description
    End Select
    Exit Sub
End Sub

Public Sub AddColumnToGraph(ByVal intGridColumn As Integer)
'Created by Robin G Brown, 20/8/97
'Insert a single column of data into the graph
'Function Declares
Const conSub = "AddColumnToGraph"
Dim lngGraphRowCounter As Long
Dim intLastColumn As Integer
    'Error Trap
    On Error GoTo AddColumnToGraph_ErrorHandler
    With graGraphControl
        intLastColumn = .DataGrid.ColumnCount
        'Fill in labels for this column
        Call .DataGrid.InsertColumns(intLastColumn + 1, 1)
        .DataGrid.ColumnLabel(intLastColumn + 1, 1) = _
            Me.Name & conColon & vfaFlexArray.TextMatrix(0, intGridColumn) _
            & conColon & vfaFlexArray.TextMatrix(1, intGridColumn)
        'Fill in data values
        For lngGraphRowCounter = 1 To vfaFlexArray.Rows - conNumFixedRows
            If vfaFlexArray.TextMatrix(lngGraphRowCounter + 1, intGridColumn) <> "" Then
                Call .DataGrid.SetData _
                    (lngGraphRowCounter, intLastColumn + 1, _
                    ConvertStringToDouble(vfaFlexArray.TextMatrix(lngGraphRowCounter + 1, intGridColumn)), False)
            Else
                Call .DataGrid.SetData(lngGraphRowCounter, intLastColumn + 1, 0#, True)
            End If
        Next
    End With
Exit Sub
'Error Handler
AddColumnToGraph_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conClass & ":" & conSub, "Unexpected error : " & Err.Description
    End Select
    Exit Sub
End Sub

Public Function GetAccountFromChartColumn(ByVal intColumn As Integer) As Integer
'Created by Robin G Brown, 13/5/97
'Return which account is in a given chart 'column', have to identify by the name in the label
'Function Declares
Dim intAccountCounter As Integer
Dim strColumnLabel As String
    'Error Trap
    On Error Resume Next
    'intColumn = intColumn - (conNumFixedCols - 1)
    'GetAccountFromColumn = (intColumn Mod conAccounts)
    With graGraphControl
        strColumnLabel = .DataGrid.ColumnLabel(intColumn, 1)
        For intAccountCounter = 1 To conAccounts
            If InStr(strColumnLabel, accAccounts(intAccountCounter).Name) > 0 Then
                'Match found
                GetAccountFromChartColumn = intAccountCounter
                Exit Function
            End If
        Next
    End With
    'Match not found!
    GetAccountFromChartColumn = intAccountCounter = -1
End Function

Public Sub ShowDisplaySettings()
'Created by Robin G Brown, 19/8/97
'Show the display settings form
'Sub Declares
Const conSub = "ShowDisplaySettings"
Dim frmSettingsForm As frmDisplaySettings
    'Error Trap
    On Error Resume Next
    'Show the settings form
    Set frmSettingsForm = New frmDisplaySettings
    Set frmSettingsForm.DataArray = Me
    frmSettingsForm.Show vbModal
    Set frmSettingsForm = Nothing
    'Redisplay the current grid
    If vfaFlexArray.TextMatrix(2, 2) <> "" Then
        Me.FillFlexArrayWithMonthlyData
    ElseIf vfaFlexArray.TextMatrix(2, 1) <> "" Then
        Me.FillFlexArrayWithQuarterlyData
    Else
        Me.FillFlexArrayWithYearlyData
    End If
End Sub

Public Sub FillFlexArrayWithMonthlyData()
'Created by Robin G Brown, 19/8/97
'Fill the given grid with data from the arrays _
    colouring, etc. as appropriate
'Show all monthly values
'Sub Declares
Const conSub = "FillFlexArrayWithMonthlyData"
Dim intDataCol As Integer
Dim intDataRow As Integer
    'Error Trap
    On Error GoTo FillFlexArrayWithMonthlyData_ErrorHandler
    Screen.MousePointer = vbHourglass
    Call SetFlexArrayDefaults
    With vfaFlexArray
        'Set the number of rows
        .Rows = (intYears * 12) + conNumFixedRows
        'Write the row headers, including spaces for subtotals
        For lngRowCounter = 1 To intYears
            For intMonthcounter = 1 To conMonths
                intDataRow = ((lngRowCounter - 1) * conMonths) + intMonthcounter
                .TextMatrix(intDataRow + 1, 0) = strRowHeaders(intDataRow, 0)
                .TextMatrix(intDataRow + 1, 1) = strRowHeaders(intDataRow, 1)
                .TextMatrix(intDataRow + 1, 2) = strRowHeaders(intDataRow, 2)
            Next
        Next
        'Set the number of fixed cols
        .FixedCols = conNumFixedCols
        'Now fill in the data, column by column
        For lngColCounter = 1 To intUWYears
            For intAcCounter = 1 To conAccounts
                intDataCol = ((lngColCounter - 1) * conAccounts) + intAcCounter
                'Write the data, row by row
                For lngRowCounter = 1 To intYears
                    For intMonthcounter = 1 To conMonths
                        intDataRow = ((lngRowCounter - 1) * conMonths) + intMonthcounter
                        Call WriteToGrid(curData(intDataRow, intDataCol), intDataCol, intDataRow + 1, intAcCounter, lngColCounter)
                    Next
                Next
            Next
        Next
        'Hide unwanted accounts and resize other columns
        Call HideSizeColumns
        .Redraw = True
    End With
    Screen.MousePointer = vbDefault
Exit Sub
'Error Handler
FillFlexArrayWithMonthlyData_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conClass & ":" & conSub, "Unexpected error : " & Err.Description
    End Select
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Public Sub FillFlexArrayWithQuarterlyData()
'Created by Robin G Brown, 19/8/97
'Fill the given grid with data from the arrays _
    colouring, etc. as appropriate
'Show all quarterly values
'Sub Declares
Const conSub = "FillFlexArrayWithQuarterlyData"
Dim intDataCol As Integer
Dim intDataRow As Integer
Dim intGridRow As Integer
Dim curCurrentData As Currency
    'Error Trap
    On Error GoTo FillFlexArrayWithQuarterlyData_ErrorHandler
    Screen.MousePointer = vbHourglass
    Call SetFlexArrayDefaults
    With vfaFlexArray
        'Set the number of rows
        .Rows = (intYears * 4) + conNumFixedRows
        'Write the row headers, including spaces for subtotals
        For lngRowCounter = 1 To intYears
            For intMonthcounter = 1 To 4
                intDataRow = ((lngRowCounter - 1) * conMonths) + (intMonthcounter * 3)
                intGridRow = ((lngRowCounter - 1) * conQuarters) + intMonthcounter
                .TextMatrix(intGridRow + 1, 0) = strRowHeaders(intDataRow, 0)
                .TextMatrix(intGridRow + 1, 1) = strRowHeaders(intDataRow, 1)
                .TextMatrix(intGridRow + 1, 2) = "" 'strRowHeaders(intDataRow, 2)
            Next
        Next
        'Set the number of fixed cols
        .FixedCols = conNumFixedCols
        'Now fill in the data, column by column
        For lngColCounter = 1 To intUWYears
            For intAcCounter = 1 To conAccounts
                intDataCol = ((lngColCounter - 1) * conAccounts) + intAcCounter
                'Write the data, row by row
                For lngRowCounter = 1 To intYears
                    For intMonthcounter = 1 To 4
                        intGridRow = ((lngRowCounter - 1) * conQuarters) + intMonthcounter
                        intDataRow = ((lngRowCounter - 1) * 4) + intMonthcounter
                        curCurrentData = SumQuarter(lngRowCounter, intMonthcounter, intDataCol)
                        Call WriteToGrid(curCurrentData, intDataCol, intGridRow + 1, intAcCounter, lngColCounter)
                    Next
                Next
            Next
        Next
        'Hide unwanted accounts and resize other columns
        Call HideSizeColumns
        'Hide the month col - RGB/17/9/97
        .ColWidth(2) = 0
        .Redraw = True
    End With
    Screen.MousePointer = vbDefault
Exit Sub
'Error Handler
FillFlexArrayWithQuarterlyData_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conClass & ":" & conSub, "Unexpected error : " & Err.Description
    End Select
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Public Sub FillFlexArrayWithYearlyData()
'Created by Robin G Brown, 19/8/97
'Fill the given grid with data from the arrays _
    colouring, etc. as appropriate
'Show all yearly values
'Sub Declares
Const conSub = "FillFlexArrayWithYearlyData"
Dim intDataCol As Integer
Dim intDataRow As Integer
Dim intGridRow As Integer
Dim curCurrentData As Currency
    'Error Trap
    On Error GoTo FillFlexArrayWithYearlyData_ErrorHandler
    Screen.MousePointer = vbHourglass
    Call SetFlexArrayDefaults
    With vfaFlexArray
        'Set the number of rows
        .Rows = intYears + conNumFixedRows
        'Write the row headers, including spaces for subtotals
        For lngRowCounter = 1 To intYears
            intDataRow = ((lngRowCounter - 1) * conMonths) + 1
            intGridRow = lngRowCounter + conNumFixedRows - 1
            .TextMatrix(intGridRow, 0) = strRowHeaders(intDataRow, 0)
            .TextMatrix(intGridRow, 1) = ""  'strRowHeaders(intDataRow, 1)
            .TextMatrix(intGridRow, 2) = ""  'strRowHeaders(intDataRow, 2)
        Next
        'Set the number of fixed cols
        .FixedCols = conNumFixedCols
        'Now fill in the data, column by column
        For lngColCounter = 1 To intUWYears
            For intAcCounter = 1 To conAccounts
                intDataCol = ((lngColCounter - 1) * conAccounts) + intAcCounter
                'Write the data, row by row
                For lngRowCounter = 1 To intYears
                    intGridRow = lngRowCounter + conNumFixedRows - 1 '((lngRowCounter - 1) * conMonths) + 1
                    intDataRow = ((lngRowCounter - 1) * 4) + 1
                    curCurrentData = SumYear(lngRowCounter, intDataCol)
                    Call WriteToGrid(curCurrentData, intDataCol, intGridRow, intAcCounter, lngColCounter)
                Next
            Next
        Next
        'Hide unwanted accounts and resize other columns
        Call HideSizeColumns
        'Hide the quarter and year cols - RGB/17/9/97
        .ColWidth(1) = 0
        .ColWidth(2) = 0
        .Redraw = True
    End With
    Screen.MousePointer = vbDefault
Exit Sub
'Error Handler
FillFlexArrayWithYearlyData_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conClass & ":" & conSub, "Unexpected error : " & Err.Description
    End Select
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

'-----------------------------------------------------------------------------
'   Special Functions and Methods
'-----------------------------------------------------------------------------

Public Sub CreateData(ByVal intStartYear As Integer)
'This sub should be private or removed for 'release'
'Created by Robin G Brown, 2/5/97
'Fill curData() with bogus data and initialise accAccounts
'Function Declares
Const conFunction = "CreateData"
Dim sngRandom As Single
Dim sngValue As Single
Dim sngPrevious As Single
Dim sngDivisor As Single
    'Error Trap
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    'Set up the sizing variables
    intYears = CInt(Format$(Now, "yyyy")) - intStartYear + 1
    intUWYears = intYears
    Randomize
    'Column Headers
    ReDim strColumnHeaders(0 To 1, 1 To intUWYears * conAccounts)
    For lngColCounter = 1 To intUWYears
        For intAcCounter = 1 To conAccounts
            strColumnHeaders(0, (lngColCounter - 1) * conAccounts + (intAcCounter - 1) + 1) = CStr(intStartYear + lngColCounter - 1)
            strColumnHeaders(1, (lngColCounter - 1) * conAccounts + (intAcCounter - 1) + 1) = DefaultAccountName(intAcCounter)
        Next
    Next
    'Row Headers
    ReDim strRowHeaders(1 To intYears * conMonths, 0 To 2)
    For lngRowCounter = 1 To intYears
        For intMonthcounter = 1 To conMonths
            strRowHeaders((lngRowCounter - 1) * conMonths + (intMonthcounter - 1) + 1, 0) = CStr(lngRowCounter)
            strRowHeaders((lngRowCounter - 1) * conMonths + (intMonthcounter - 1) + 1, 1) = DefaultQuarterName(intMonthcounter)
            strRowHeaders((lngRowCounter - 1) * conMonths + (intMonthcounter - 1) + 1, 2) = DefaultPeriodName(intMonthcounter)
        Next
    Next
    'Currency data
    ReDim curData(1 To intYears * conMonths, 1 To intUWYears * conAccounts)
    For lngColCounter = 1 To intUWYears
        For intAcCounter = 1 To conAccounts
            sngRandom = Rnd * 10
            For lngRowCounter = 1 To intYears
                For intMonthcounter = 1 To conMonths
                    If lngRowCounter <= (intYears + 1 - lngColCounter) Then
                        If lngRowCounter = 1 And intMonthcounter = 1 Then
                            sngValue = sngRandom
                            sngPrevious = sngValue
                            sngDivisor = -1
                        Else
                            If sngDivisor < 1 Then sngDivisor = 0.5
                            sngValue = sngPrevious + (sngRandom / sngDivisor)
                            sngPrevious = sngValue
                            sngDivisor = sngDivisor + 1
                        End If
                        curData((lngRowCounter - 1) * conMonths + (intMonthcounter - 1) + 1, (lngColCounter - 1) * conAccounts + (intAcCounter - 1) + 1) _
                            = sngValue
                    Else
                        curData((lngRowCounter - 1) * conMonths + (intMonthcounter - 1) + 1, (lngColCounter - 1) * conAccounts + (intAcCounter - 1) + 1) _
                            = 0#
                    End If
                Next
            Next
        Next
    Next
    'Set up accAccounts
    Call InitialiseFormat
    Screen.MousePointer = vbDefault
End Sub

'-----------------------------------------------------------------------------
'   Private Functions and Methods
'-----------------------------------------------------------------------------

Private Sub InitialiseData()
'Created by Robin G Brown, 7/5/97
'Initialise the data array
'Sub Declares
    'Error Trap
    On Error Resume Next
    'Column Headers
    ReDim strColumnHeaders(0 To 1, 1 To intYears * conAccounts)
    For lngColCounter = 1 To intUWYears
        For intAcCounter = 1 To conAccounts
            strColumnHeaders(0, (lngColCounter - 1) * conAccounts + (intAcCounter - 1) + 1) = ""
            strColumnHeaders(1, (lngColCounter - 1) * conAccounts + (intAcCounter - 1) + 1) = ""
        Next
    Next
    'Row Headers
    ReDim strRowHeaders(1 To intYears * conMonths, 0 To 2)
    For lngRowCounter = 1 To intYears
        For intMonthcounter = 1 To conMonths
            strRowHeaders((lngRowCounter - 1) * conMonths + (intMonthcounter - 1) + 1, 0) = ""
            strRowHeaders((lngRowCounter - 1) * conMonths + (intMonthcounter - 1) + 1, 1) = ""
            strRowHeaders((lngRowCounter - 1) * conMonths + (intMonthcounter - 1) + 1, 2) = ""
        Next
    Next
    'Currency data
    ReDim curData(1 To intYears, 1 To intUWYears * conAccounts)
    For lngRowCounter = 1 To intYears
        For lngColCounter = 1 To intUWYears * conAccounts
            curData(lngRowCounter, lngColCounter) = 0#
        Next
    Next
End Sub

Private Sub InitialiseFormat()
'Created by Robin G Brown, 7/5/97
'Initialise the displaycolumn array
'Sub Declares
    'Error Trap
    On Error Resume Next
    ReDim accAccounts(1 To conAccounts)
    For intAcCounter = 1 To conAccounts
        accAccounts(intAcCounter).Name = DefaultAccountName(intAcCounter)
        If InStr(accAccounts(intAcCounter).Name, "%") > 0 Then
            accAccounts(intAcCounter).Display = False
            accAccounts(intAcCounter).Mask = "0\%"
        Else
            accAccounts(intAcCounter).Display = True
            accAccounts(intAcCounter).Mask = "0"
        End If
    Next
End Sub

Private Function DefaultAccountName(ByVal intAccount As Integer) As String
'Created by Robin G Brown, 12/5/97
'Return a string describing the account
'Function Declares
    'Error Trap
    On Error Resume Next
    Select Case intAccount
    Case 1
        DefaultAccountName = "PREM"
    Case 2
        DefaultAccountName = "PAID"
    Case 3
        DefaultAccountName = "PAID%"
    Case 4
        DefaultAccountName = "O/S"
    Case 5
        DefaultAccountName = "O/S%"
    Case 6
        DefaultAccountName = "INC"
    Case 7
        DefaultAccountName = "INC%"
    Case 8
        DefaultAccountName = "IBNR"
    Case 9
        DefaultAccountName = "IBNR%"
    Case Else
        DefaultAccountName = "Unknown"
    End Select
End Function

Private Function DefaultPeriodName(ByVal intAccount As Integer) As String
'Created by Robin G Brown, 12/5/97
'Return a string describing the account
'Function Declares
    'Error Trap
    On Error Resume Next
    Select Case intAccount
    Case 1
        DefaultPeriodName = "Jan"
    Case 2
        DefaultPeriodName = "Feb"
    Case 3
        DefaultPeriodName = "Mar"
    Case 4
        DefaultPeriodName = "Apr"
    Case 5
        DefaultPeriodName = "May"
    Case 6
        DefaultPeriodName = "Jun"
    Case 7
        DefaultPeriodName = "Jul"
    Case 8
        DefaultPeriodName = "Aug"
    Case 9
        DefaultPeriodName = "Sep"
    Case 10
        DefaultPeriodName = "Oct"
    Case 11
        DefaultPeriodName = "Nov"
    Case 12
        DefaultPeriodName = "Dec"
    Case Else
        DefaultPeriodName = "Unknown"
    End Select
End Function

Private Function DefaultQuarterName(ByVal intAccount As Integer) As String
'Created by Robin G Brown, 12/8/97
'Return a string describing the quarter
'Function Declares
    'Error Trap
    On Error Resume Next
    Select Case intAccount
    Case 1, 2, 3
        DefaultQuarterName = "Q1"
    Case 4, 5, 6
        DefaultQuarterName = "Q2"
    Case 7, 8, 9
        DefaultQuarterName = "Q3"
    Case 10, 11, 12
        DefaultQuarterName = "Q4"
    Case Else
        DefaultQuarterName = "Unknown"
    End Select
End Function

Private Function SumYear(ByVal intYear As Integer, ByVal intDataColumn As Integer) As Currency
'Created by Robin G Brown, 21/5/97
'Returns the sum of the periods for the given intYear and intDataColumn
'Function Declares
Dim intSumCounter As Integer
Dim curSum As Currency
    'Error Trap
    On Error GoTo SumYear_ErrorHandler
    curSum = 0#
    For intSumCounter = ((intYear - 1) * conMonths) + 1 To intYear * conMonths
        curSum = curSum + curData(intSumCounter, intDataColumn)
    Next
    SumYear = curSum
Exit Function
'Error Handler
SumYear_ErrorHandler:
    SumYear = 0#
    Exit Function
End Function

Private Function SumQuarter(ByVal intYear As Integer, ByVal intQuarter As Integer, ByVal intDataColumn As Integer) As Currency
'Created by Robin G Brown, 21/5/97
'Returns the sum of the periods for the given intYear, intQuarter and intDataColumn
'Function Declares
Dim intSumCounter As Integer
Dim curSum As Currency
Dim intStartMonth As Integer
Dim intEndMonth As Integer
    'Error Trap
    On Error GoTo SumQuarter_ErrorHandler
    curSum = 0#
    intStartMonth = ((intYear - 1) * conMonths) + 1
    intEndMonth = intYear * conMonths
    Select Case intQuarter
    Case 1
        intStartMonth = intStartMonth
        intEndMonth = intEndMonth - 9
        For intSumCounter = intStartMonth To intEndMonth
            curSum = curSum + curData(intSumCounter, intDataColumn)
        Next
    Case 2
        intStartMonth = intStartMonth + 3
        intEndMonth = intEndMonth - 6
        For intSumCounter = intStartMonth To intEndMonth
            curSum = curSum + curData(intSumCounter, intDataColumn)
        Next
    Case 3
        intStartMonth = intStartMonth + 6
        intEndMonth = intEndMonth - 3
        For intSumCounter = intStartMonth To intEndMonth
            curSum = curSum + curData(intSumCounter, intDataColumn)
        Next
    Case 4
        intStartMonth = intStartMonth + 9
        intEndMonth = intEndMonth
        For intSumCounter = intStartMonth To intEndMonth
            curSum = curSum + curData(intSumCounter, intDataColumn)
        Next
    Case Else
        'Error
        SumQuarter = 0#
    End Select
    SumQuarter = curSum
Exit Function
'Error Handler
SumQuarter_ErrorHandler:
    SumQuarter = 0#
    Exit Function
End Function

Private Sub SetFlexArrayDefaults()
'Created by Robin G Brown, 19/8/97
'Set up vfaFlexArray
'Sub Declares
Dim intDataCol As Integer
    'Error Trap
    On Error Resume Next
    With vfaFlexArray
        .FillStyle = flexFillRepeat
        .Redraw = False
        'Allow rows and cols only to merge
        .MergeCells = flexMergeRestrictAll
        'Set only the top row and first 2 cols to merge
        .MergeRow(0) = True
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        'Set the number of fixed cols and rows
        .FixedCols = 0
        .FixedRows = 0
        'Set the number of cols and a default value for rows
        .Cols = (intUWYears * conAccounts) + conNumFixedCols
        .Rows = 3
        'Set the number of fixed cols and rows
        .FixedCols = 1
        .FixedRows = conNumFixedRows
        'Write the col headers
        For lngColCounter = 1 To intUWYears
            'For each year add a number of accounts
            For intAcCounter = 1 To conAccounts
                intDataCol = ((lngColCounter - 1) * conAccounts) + intAcCounter
                .TextMatrix(0, intDataCol + conNumFixedCols - 1) = strColumnHeaders(0, intDataCol)
                .TextMatrix(1, intDataCol + conNumFixedCols - 1) = strColumnHeaders(1, intDataCol)
            Next
        Next
        .TextMatrix(0, 0) = "Year"
        .TextMatrix(0, 1) = "Quarter"
        .TextMatrix(0, 2) = "Month"
        .TextMatrix(1, 0) = "Year"
        .TextMatrix(1, 1) = "Quarter"
        .TextMatrix(1, 2) = "Month"
    End With
End Sub

Private Sub WriteToGrid(ByVal curValue As Currency, ByVal intDataCol As Integer, ByVal intGridRow As Integer, ByVal intAccount As Integer, ByVal intCurColumn As Integer)
'Created by Robin G Brown, 19/8/97
'Write the data to the grid, applying any format required
'Sub Declares
    'Error Trap
    On Error Resume Next
    With vfaFlexArray
        If curValue <> 0# Then
            .TextMatrix(intGridRow + conNumFixedRows - 2, intDataCol + conNumFixedCols - 1) = _
                Format(curValue, accAccounts(intAccount).Mask)
            'Format the cells
            .Col = intDataCol + conNumFixedCols - 1
            .Row = intGridRow + conNumFixedRows - 2
            If intCurColumn Mod 2 = 0 Then
                .CellBackColor = &HE0E0E0
            Else
                .CellBackColor = vbWhite
            End If
        Else
            'Show a blank if the value is 0
            .TextMatrix(intGridRow + conNumFixedRows - 2, intDataCol + conNumFixedCols - 1) = ""
            .Col = intDataCol + conNumFixedCols - 1
            .Row = intGridRow + conNumFixedRows - 2
            .CellBackColor = .BackColorBkg
        End If
    End With
End Sub

Private Sub HideSizeColumns()
'Created by Robin G Brown, 19/8/97
'Hide unwanted accounts and resize other columns
'Sub Declares
Dim intDataCol As Integer
    'Error Trap
    On Error Resume Next
    With vfaFlexArray
        .AutoSize 0, 2, False
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .AutoSize conNumFixedCols, .Cols - 1, True
        For lngColCounter = 1 To intUWYears
            For intAcCounter = 1 To conAccounts
                intDataCol = ((lngColCounter - 1) * conAccounts) + intAcCounter
                .ColAlignment(intDataCol + conNumFixedCols - 1) = flexAlignCenterCenter
                If accAccounts(intAcCounter).Display = True Then
                    'Do nothing
                Else
                    .ColWidth(intDataCol + conNumFixedCols - 1) = 0
                End If
            Next
        Next
    End With
End Sub

Private Sub PrepareChart()
'Created by Robin G Brown, 21/8/97
'Set up the chart for graphing data
'Sub Declares
    'Error Trap
    On Error Resume Next
    With graGraphControl
        'Set the size of the datagrid first
        Call .DataGrid.SetSize(2, 1, vfaFlexArray.Rows - 2, 0)
        'Fill in labels for rows
        For lngRowCounter = 1 To vfaFlexArray.Rows - 2
            .DataGrid.RowLabel(lngRowCounter, 1) = _
                vfaFlexArray.TextMatrix(lngRowCounter + 1, 0)
            .DataGrid.RowLabel(lngRowCounter, 2) = _
                vfaFlexArray.TextMatrix(lngRowCounter + 1, 1)
        Next
    End With
End Sub

'-----------------------------------------------------------------------------
'   Properties
'-----------------------------------------------------------------------------

Property Get DataRows() As Long
'Created by Robin G Brown, 28/8/97
'Return the number of data rows
'Property Declares
Const conProperty = "DataRows"
'Error Trap
    On Error Resume Next
    DataRows = vfaFlexArray.Rows - vfaFlexArray.FixedRows
End Property

Property Get GridMouseCol() As Long
'Created by Robin G Brown, 20/8/97
'Property Declares
Const conProperty = "GridMouseCol"
'Error Trap
    On Error Resume Next
    GridMouseCol = grmGridMouse.Col
End Property

Property Let GridMouseCol(ByVal lngNewGridMouseCol As Long)
'Created by Robin G Brown, 20/8/97
'Sub Declares
Const conProperty = "GridMouseCol"
    'Error Trap
    On Error Resume Next
    If lngNewGridMouseCol >= 0 And lngNewGridMouseCol <= vfaFlexArray.Cols Then
        grmGridMouse.Col = lngNewGridMouseCol
    Else
        grmGridMouse.Col = -1
        Err.Raise Err.Number, conClass & ":" & conProperty, "Unexpected error : " & Err.Description
    End If
End Property

Property Get GridMouseRow() As Long
'Created by Robin G Brown, 20/8/97
'Property Declares
Const conProperty = "GridMouseRow"
'Error Trap
    On Error Resume Next
    GridMouseRow = grmGridMouse.Row
End Property

Property Let GridMouseRow(ByVal lngNewGridMouseRow As Long)
'Created by Robin G Brown, 20/8/97
'Sub Declares
Const conProperty = "GridMouseRow"
    'Error Trap
    On Error Resume Next
    If lngNewGridMouseRow >= 0 And lngNewGridMouseRow <= vfaFlexArray.Rows Then
        grmGridMouse.Row = lngNewGridMouseRow
    Else
        grmGridMouse.Row = -1
        Err.Raise Err.Number, conClass & ":" & conProperty, "Unexpected error : " & Err.Description
    End If
End Property

Property Get DisplayColumn(ByVal intColumn As Integer) As Boolean
'Created by Robin G Brown, 6/5/97
'Return whether or not to display a column, _
    based on its absolute position, _
    assuming that the first column is no. 1 _
    and that there are 2 fixed columns
'Function Declares
Dim intAccount As Integer
    'Error Trap
    On Error Resume Next
    'Adjust for the 2 fixed columns
    intColumn = intColumn - 1
    intAccount = (intColumn Mod conAccounts)
    If intAccount = 0 Then
        intAccount = conAccounts
    End If
    DisplayColumn = accAccounts(intAccount).Display
End Property

Property Get AccountName(ByVal intAccount As Integer) As String
'Created by Robin G Brown, 29/8/97
'Return the name of the account in strColumnheaders
'Property Declares
Const conProperty = "AccountName"
'Error Trap
    On Error Resume Next
    AccountName = accAccounts(intAccount).Name
End Property

Property Let AccountName(ByVal intAccount As Integer, ByVal strNewName As String)
'Created by Robin G Brown, 14/5/97
'Set the name of the account in strColumnheaders
'Sub Declares
    'Error Trap
    On Error Resume Next
    'Set the name
    accAccounts(intAccount).Name = strNewName
    'Fill in strColumnHeaders
    For lngColCounter = 1 To intUWYears
        'strColumnHeaders(0, (lngColCounter - 1) * conAccounts + (intAccount - 1) + 1) = ""
        strColumnHeaders(1, (lngColCounter - 1) * conAccounts + (intAccount - 1) + 1) = strNewName
    Next
End Property

Property Get VisibleAccounts() As Integer
'Created by Robin G Brown, 7/5/97
'Returns the number of visible accounts
'Function Declares
Dim intVisible As Integer
    'Error Trap
    On Error Resume Next
    intVisible = 0
    For intAcCounter = 1 To conAccounts
        If accAccounts(intAcCounter).Display = True Then
            intVisible = intVisible + 1
        End If
    Next
    VisibleAccounts = intVisible
End Property

Property Get DisplayAccount(ByVal intAccount As Integer) As Boolean
'Created by Robin G Brown, 6/5/97
'Return wether or not to display an account
'Function Declares
    'Error Trap
    On Error Resume Next
    DisplayAccount = accAccounts(intAccount).Display
End Property

Property Get AccountMask(ByVal intAccount As Integer) As String
'Created by Robin G Brown, 6/5/97
'Return the display colour of an account
'Function Declares
    'Error Trap
    On Error Resume Next
    AccountMask = accAccounts(intAccount).Mask
End Property

Property Let AccountMask(ByVal intAccount As Integer, ByVal strNewMask As String)
'Created by Robin G Brown, 6/5/97
'Set the mask of an account
'Function Declares
    'Error Trap
    On Error Resume Next
    accAccounts(intAccount).Mask = strNewMask
End Property

Property Get FlexArray() As Object
'Created by Robin G Brown, 12/8/97
'Set vfaFlexArray to a VSFlexArray control
'Sub Declares
    'Error Trap
    On Error Resume Next
    Set FlexArray = vfaFlexArray
End Property

Property Set FlexArray(ByRef vfaNewFlexArray As Object)
'Created by Robin G Brown, 12/8/97
'Set vfaFlexArray to a VSFlexArray control
'Sub Declares
    'Error Trap
    On Error Resume Next
    If TypeOf vfaNewFlexArray Is vsFlexArray Then
        Set vfaFlexArray = vfaNewFlexArray
    Else
        Err.Raise conInvalidObjectError, conClass, "Invalid object"
    End If
End Property

Property Set Chart(ByRef graNewGraphControl As Object)
'Created by Robin G Brown, 12/8/97
'Set graGraphControl to an MSChart control
'Sub Declares
    'Error Trap
    On Error Resume Next
    If TypeOf graNewGraphControl Is MSChart Then
        Set graGraphControl = graNewGraphControl
    Else
        Err.Raise conInvalidObjectError, conClass, "Invalid object"
    End If
End Property

Property Get Accounts() As Integer
'Created by Robin G Brown, 6/5/97
'Return the number of accounts
'Function Declares
    'Error Trap
    On Error Resume Next
    Accounts = conAccounts
End Property

Property Let Accounts(ByVal intNumberOfAccounts As Integer)
'DISABLED - you are not allowed to change the number of accounts - RGB/14/5/97
Const conSub = "Accounts"
    Err.Raise vbObjectError, conClass & ":" & conSub, "Number of accounts may not be altered."
End Property

Property Get Name() As String
'Created by Robin G Brown, 14/5/97
'Return the name of this instance
'Function Declares
    'Error Trap
    On Error Resume Next
    Name = strInstance
End Property

Property Let Name(ByVal strNewName As String)
'Created by Robin G Brown, 14/5/97
'Set the name of this instance
'Sub Declares
    'Error Trap
    On Error Resume Next
    strInstance = strNewName
End Property

Property Get UWYears() As Integer
'Created by Robin G Brown, 6/5/97
'Return the Years value
'Function Declares
    'Error Trap
    On Error Resume Next
    UWYears = intUWYears
End Property

Property Let UWYears(ByVal intNewValue As Integer)
'Created by Robin G Brown, 6/5/97
'Set the years value, and reinitialise
'Sub Declares
    'Error Trap
    On Error Resume Next
    If intNewValue < 0 Then
        Err.Raise conBadYearValueError, conClass, "Years value must be greater than 0"
    Else
        intUWYears = intNewValue
    End If
    Call InitialiseData
    Call InitialiseFormat
End Property

Property Get Years() As Integer
'Created by Robin G Brown, 6/5/97
'Return the Years value
'Function Declares
    'Error Trap
    On Error Resume Next
    Years = intYears
End Property

Property Let Years(ByVal intNewValue As Integer)
'Created by Robin G Brown, 6/5/97
'Set the years value, and reinitialise
'Sub Declares
    'Error Trap
    On Error Resume Next
    If intNewValue < 0 Then
        Err.Raise conBadYearValueError, conClass, "Years value must be greater than 0"
    Else
        intYears = intNewValue
    End If
    Call InitialiseData
    Call InitialiseFormat
End Property

Private Sub Class_Terminate()
'Created by Robin G Brown, 6/5/97
'Default Behaviour
'Sub Declares
    'Error Trap
    On Error Resume Next
    'Nobble the arrays
    ReDim curData(1 To 1, 1 To 1)
    ReDim strColumnHeaders(1 To 1, 1 To 1)
    ReDim strRowHeaders(1 To 1, 1 To 1)
    ReDim accAccounts(1 To 1)
    'Clear the references
    Set vfaFlexArray = Nothing
    Set graGraphControl = Nothing
End Sub



