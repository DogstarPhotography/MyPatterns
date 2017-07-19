VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Created by Robin G Brown, 6th May 1997
'-----------------------------------------------------------------------------
'   This class module contains code for _
    holding and manipulating the data for the grid
'   Use dta as a prefix for this class
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "DataArray"
'Public properties
'Private properties
Private intCounter As Integer
Private intReturn As Integer
Private varData() As Variant
Private booShowAccount() As Boolean
Private intUWYears As Integer
Private intAccounts As Integer
Private intPeriods As Integer
Private intRowCounter As Integer
Private intColCounter As Integer
Private intAcCounter As Integer
'Error Constants
Private Const errArray = 55001
Private Const errMismatch = 55002
Private Const errTooManyAccounts = 55003
Private Const errBadYearValue = 55004
'-----------------------------------------------------------------------------
'Spoiler information for DataArray
'
'   1. DataArray represents a grid with a single header row & column
'   2. The index is always (row, column)
'   3. The row headers are in position (n,0)
'   4. The column headers are in position (0,n)
'   5. Data starts in position (1,1)
'   6. Position (0,0) can be used for 'tag' information
'   7. Columns are grouped into years with a size equal to intAccounts
'   8. Total number of data columns is intUWYears * intAccounts (numbered from 1),
'       plus 1 header column (0)
'   9. Total number of data rows is equal to intPeriods (numbered from 1),
'       plus 1 header row (0)
'   10. booShowAccount is used as a 'mask' to determine
'       whether account columns should be displayed FOR EACH YEAR
'
'-----------------------------------------------------------------------------

Private Sub Class_Initialize()
'Created by Robin G Brown, 6/5/97
'Default Behaviour
'Sub Declares
    'Error Trap
    On Error Resume Next
    'Initialise the arrays and default values
    intPeriods = 15
    intUWYears = 15
    intAccounts = 3
    Call InitialiseData
    Call InitialiseAccount
End Sub

Private Sub InitialiseData()
'Created by Robin G Brown, 7/5/97
'Initialise the data array
'Sub Declares
    'Error Trap
    On Error Resume Next
    ReDim varData(0 To intPeriods, 0 To intPeriods * intAccounts)
    'Column Headers
    For intCounter = 1 To intUWYears
        varData(intCounter, 0) = ""
        For intAcCounter = 1 To intAccounts
            varData(0, (intCounter - 1) * intAccounts + (intAcCounter - 1) + 1) = ""
        Next
    Next
    For intColCounter = 1 To intPeriods
        For intAcCounter = 1 To intAccounts
            For intRowCounter = 1 To intPeriods
                varData(intRowCounter, (intColCounter - 1) * intAccounts + (intAcCounter - 1) + 1) = 0#
            Next
        Next
    Next
End Sub

Private Sub InitialiseAccount()
'Created by Robin G Brown, 7/5/97
'Initialise the displaycolumn array
'Sub Declares
    'Error Trap
    On Error Resume Next
    ReDim booShowAccount(1 To intAccounts)
    For intCounter = 1 To intAccounts
        booShowAccount(intCounter) = True
    Next
End Sub

Private Sub Class_Terminate()
'Created by Robin G Brown, 6/5/97
'Default Behaviour
'Sub Declares
    'Error Trap
    On Error Resume Next
    'Nobble the array
    ReDim varData(1 To 1)
    ReDim booShowAccount(1 To 1)
End Sub

Property Get Periods() As Integer
'Created by Robin G Brown, 6/5/97
'Return the Years value
'Function Declares
    'Error Trap
    On Error Resume Next
    Periods = intPeriods
End Property

Property Let Periods(ByVal intNewValue As Integer)
'Created by Robin G Brown, 6/5/97
'Set the years value, and reinitialise
'Sub Declares
    'Error Trap
    On Error Resume Next
    If intNewValue < 0 Then
        ThrowError errBadYearValue
    Else
        intPeriods = intNewValue
    End If
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
        ThrowError errBadYearValue
    Else
        intUWYears = intNewValue
    End If
End Property

Property Get Accounts() As Integer
'Created by Robin G Brown, 6/5/97
'Return the number of accounts
'Function Declares
    'Error Trap
    On Error Resume Next
    Accounts = intAccounts
End Property

Public Function VisibleAccounts() As Integer
'Created by Robin G Brown, 7/5/97
'Returns the number of visible accounts
'Function Declares
Dim intVisible As Integer
    'Error Trap
    On Error Resume Next
    intVisible = 0
    For intCounter = 1 To intAccounts
        If booShowAccount(intCounter) = True Then
            intVisible = intVisible + 1
        End If
    Next
    VisibleAccounts = intVisible
End Function

Property Let Accounts(ByVal intNewValue As Integer)
'Created by Robin G Brown, 6/5/97
'Set the number of accounts, and reinitialise
'Sub Declares
    'Error Trap
    On Error Resume Next
    If intNewValue < 0 Or intNewValue > 6 Then
        ThrowError errTooManyAccounts
    Else
        intAccounts = intNewValue
        ReDim booShowAccount(1 To intAccounts)
        For intCounter = 1 To intAccounts
            booShowAccount(intCounter) = True
        Next
    End If
End Property

Public Function LoadData(ByRef varArray() As Variant) As Boolean
'Created by Robin G Brown, 2/5/97
'Fill varData() from an array of data - varArray()
'Function Declares
Const conFunction = "LoadData"
Dim intLBoundD1 As Integer
Dim intUBoundD1 As Integer
Dim intDimension1Counter As Integer
Dim intLBoundD2 As Integer
Dim intUBoundD2 As Integer
Dim intDimension2Counter As Integer
    'Error Trap
    On Error GoTo LoadData_ErrorHandler
    'Read the size of the array
    intLBoundD1 = LBound(varArray, 1)
    intUBoundD1 = UBound(varArray, 1)
    intLBoundD2 = LBound(varArray, 1)
    intUBoundD2 = UBound(varArray, 1)
    'REDO THIS CODE! - RGB/12/5/97
    'Now we know how big it is we can load the data into our own array
    ReDim varData(intLBoundD1 To intUBoundD1, intLBoundD2 To intUBoundD2)
    For intDimension1Counter = intLBoundD2 To intUBoundD1
        For intDimension2Counter = intLBoundD2 To intUBoundD2
            varData(intDimension1Counter, intDimension2Counter) = _
                varArray(intDimension1Counter, intDimension2Counter)
        Next
    Next
    LoadData = True
Exit Function
'Error Handler
LoadData_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conFunction
    End Select
    ReDim varData(1 To 1, 1 To 1)
    LoadData = False
    Exit Function
End Function

Public Sub CopyData(ByRef varArray() As Variant)
'Created by Robin G Brown, 2/5/97
'Copy varData into varArray, resizing varArray in the process
'Function Declares
Const conSub = "CopyData"
Dim intLBoundD1 As Integer
Dim intUBoundD1 As Integer
Dim intDimension1Counter As Integer
Dim intLBoundD2 As Integer
Dim intUBoundD2 As Integer
Dim intDimension2Counter As Integer
    'Error Trap
    On Error GoTo CopyData_ErrorHandler
    'Read the size of the array
    intLBoundD1 = LBound(varData, 1)
    intUBoundD1 = UBound(varData, 1)
    intLBoundD2 = LBound(varData, 2)
    intUBoundD2 = UBound(varData, 2)
    'REDO THIS CODE - RGB/12/5/97
    'Now we know how big it is we can load the data into our own array
    ReDim varArray(intLBoundD1 To intUBoundD1, intLBoundD2 To intUBoundD2)
    For intDimension1Counter = intLBoundD2 To intUBoundD1
        For intDimension2Counter = intLBoundD2 To intUBoundD2
            varArray(intDimension1Counter, intDimension2Counter) = _
                varData(intDimension1Counter, intDimension2Counter)
        Next
    Next
Exit Sub
'Error Handler
CopyData_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    ReDim varArray(1 To 1, 1 To 1)
    Exit Sub
End Sub

Property Get GetElement(ByVal intD1Pos As Integer, ByVal intD2Pos As Integer) As Variant
'Created by Robin G Brown, 2/5/97
'Return the value of the given element
'Function Declares
    'Error Trap
    On Error Resume Next
    GetElement = varData(intD1Pos, intD2Pos)
    If Err.Number <> 0 Then
        GetElement = Null
        Err.Clear
    End If
End Property

Property Let SetElement(ByVal intD1Pos As Integer, ByVal intD2Pos As Integer, ByVal varValue As Variant)
'Created by Robin G Brown, 2/5/97
'Set the value of the given element
'Sub Declares
    'Error Trap
    On Error Resume Next
    varData(intD1Pos, intD2Pos) = varValue
    If Err.Number <> 0 Then
        'Raise an error??
        'ThrowError errArray
    End If
End Property

Public Function ShowAccount(ByVal intAccount As Integer, booNewDisplaySetting As Boolean)
'Created by Robin G Brown, 7/5/97
'Sets booShowAccount(intAccount) according to booNewDisplaySetting
'Sub Declares
    'Error Trap
    On Error Resume Next
    booShowAccount(intAccount) = booNewDisplaySetting
End Function

Public Function ShowAccounts(ByRef booNewDisplaySetting() As Boolean)
'Created by Robin G Brown, 7/5/97
'Sets booShowAccount according to booNewDisplaySetting, _
    which must have elements starting at 1 and intAccounts number of elements
'Sub Declares
    'Error Trap
    On Error Resume Next
    If UBound(booNewDisplaySetting, 1) <> intAccounts Then
        ThrowError errMismatch
        Exit Function
    Else
        For intCounter = 1 To intAccounts
            booShowAccount(intCounter) = booNewDisplaySetting(intCounter)
        Next
    End If
End Function

Public Function ShowAllAccounts()
'Created by Robin G Brown, 7/5/97
'Sets booShowAccount to show all account columns
'Sub Declares
    'Error Trap
    On Error Resume Next
    For intCounter = 1 To intAccounts
        booShowAccount(intCounter) = True
    Next
End Function

Public Function DisplayColumn(ByVal intcolumn As Integer) As Boolean
'Created by Robin G Brown, 6/5/97
'Return whether or not to display a column, _
    based on its absolute position, _
    assuming that the first column is no. 1
'Function Declares
Dim intAccount As Integer
    'Error Trap
    On Error Resume Next
    intAccount = (intcolumn Mod intAccounts)
    If intAccount = 0 Then
        intAccount = intAccounts
    End If
    DisplayColumn = booShowAccount(intAccount)
End Function

Public Function DisplayAccount(ByVal intAccount As Integer) As Boolean
'Created by Robin G Brown, 6/5/97
'Return wether or not to display a column
'Function Declares
    'Error Trap
    On Error Resume Next
    DisplayAccount = booShowAccount(intAccount)
End Function

Public Sub ThrowError(ByVal errNumber As Long)
'Created by Robin G Brown, 2/5/97
'Raise an error
'Sub Declares
    'Error Trap
    On Error Resume Next
    Select Case errNumber
    Case errArray
        Err.Raise errArray, conModule, "Array error"
    Case errMismatch
        Err.Raise errMismatch, conModule, "Array mismatch"
    Case errTooManyAccounts
        Err.Raise errTooManyAccounts, conModule, "Too many/Too few accounts"
    Case errBadYearValue
        Err.Raise errBadYearValue, conModule, "Years value must be greater than 0"
    Case Else
        Err.Raise 55000, conModule, "Undefined Error"
    End Select
End Sub

Public Sub CreateData()
'This sub should be private for 'release'
'Created by Robin G Brown, 2/5/97
'Fill varData() with bogus data and initialise booShowAccount
'Function Declares
Const conFunction = "CreateData"
Dim sngRandom As Single
Dim sngValue As Single
Dim sngPrevious As Single
Dim sngDivisor As Single
    'Error Trap
    On Error Resume Next
    Randomize
    ReDim varData(0 To intPeriods, 0 To intUWYears * intAccounts)
    'Column Headers
    For intCounter = 1 To intUWYears
        varData(intCounter, 0) = "Period:" & intCounter
        For intAcCounter = 1 To intAccounts
            varData(0, (intCounter - 1) * intAccounts + (intAcCounter - 1) + 1) _
                = "UWYear:" & intCounter & ",Account:" & intAcCounter
        Next
    Next
    For intColCounter = 1 To intUWYears
        For intAcCounter = 1 To intAccounts
            sngRandom = Rnd * 100
            For intRowCounter = 1 To intPeriods
                If intRowCounter <= (intPeriods + 1 - intColCounter) Then
                    If intRowCounter = 1 Then
                        sngValue = sngRandom
                        sngPrevious = sngValue
                        sngDivisor = -1
                    Else
                        If sngDivisor < 1 Then sngDivisor = 0.5
                        sngValue = sngPrevious + (sngRandom / sngDivisor)
                        sngPrevious = sngValue
                        sngDivisor = sngDivisor + 1
                    End If
                    varData(intRowCounter, (intColCounter - 1) * intAccounts + (intAcCounter - 1) + 1) _
                        = Format$(sngValue, "00.00")
                Else
                    varData(intRowCounter, (intColCounter - 1) * intAccounts + (intAcCounter - 1) + 1) _
                        = "NULL"
                End If
            Next
        Next
    Next
    'Set up booShowAccount
    Call InitialiseAccount
End Sub

Public Sub FillGrid(ByRef grdControl As MSFlexGrid)
'Created by Robin G Brown, 2/5/97
'Default behaviour
'Sub Declares
Const conSub = "FillGrid"
Dim intGridCol As Integer
Dim intDataCol As Integer
Dim intNumVisAccounts As Integer
Dim intOffset As Integer
    'Error Trap
    On Error GoTo FillGrid_ErrorHandler
    grdControl.Clear
    'Set the number of cols and rows
    grdControl.FixedCols = 0
    grdControl.FixedRows = 0
    intNumVisAccounts = VisibleAccounts
    grdControl.Cols = (intPeriods * intNumVisAccounts) + 1
    grdControl.Rows = intPeriods + 1
    'Write the col headers
    For intColCounter = 1 To intUWYears
        intOffset = 1
        For intAcCounter = 1 To intAccounts
            If DisplayAccount(intAcCounter) = True Then
                intGridCol = ((intColCounter - 1) * intNumVisAccounts) + intOffset
                intDataCol = ((intColCounter - 1) * intAccounts) + intAcCounter
                grdControl.TextMatrix(0, intGridCol) = varData(0, intDataCol)
                intOffset = intOffset + 1
            End If
        Next
    Next
    For intRowCounter = 1 To intPeriods
        'Write the row header
        grdControl.TextMatrix(intRowCounter, 0) = varData(intRowCounter, 0)
        For intColCounter = 1 To intUWYears
            intOffset = 1
            For intAcCounter = 1 To intAccounts
                'Write the data
                If DisplayAccount(intAcCounter) = True Then
                    intGridCol = ((intColCounter - 1) * intNumVisAccounts) + intOffset
                    intDataCol = ((intColCounter - 1) * intAccounts) + intAcCounter
                    If IsNumeric(varData(intRowCounter, intDataCol)) = True Then
                        grdControl.TextMatrix(intRowCounter, intGridCol) = _
                            varData(intRowCounter, intDataCol)
                        intOffset = intOffset + 1
                    End If
                End If
            Next
        Next
    Next
    'Set the number of fixed cols and rows
    grdControl.FixedCols = 1
    grdControl.FixedRows = 1
Exit Sub
'Error Handler
FillGrid_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Public Sub FillChart(ByRef graControl As MSChart)
'Created by Robin G Brown, 6/5/97
'This function copied from the Chart Control Library
'Set up a chart from the varData() array
'Sub Declares
Const conSub = "FillChart"
Dim intCols As Integer
Dim intRows As Integer
Dim intGridCol As Integer
Dim intDataCol As Integer
Dim intNumVisAccounts As Integer
Dim intOffset As Integer
    'Error Trap
    On Error GoTo FillChart_ErrorHandler
    'Set the size of the datagrid first
    intNumVisAccounts = VisibleAccounts
    intRows = intPeriods
    intCols = (intUWYears * intNumVisAccounts)
    Call graControl.DataGrid.SetSize(intRows, intCols, intRows, intCols)
    'Fill in labels for rows
    For intRowCounter = 1 To intRows
        graControl.DataGrid.RowLabel(intRowCounter, 1) = _
            CStr(varData(intRowCounter, 0))
    Next
    'Fill in labels for columns
    For intColCounter = 1 To intUWYears
        intOffset = 1
        For intAcCounter = 1 To intAccounts
            If DisplayAccount(intAcCounter) = True Then
                intGridCol = ((intColCounter - 1) * intNumVisAccounts) + intOffset
                intDataCol = ((intColCounter - 1) * intAccounts) + intAcCounter
                graControl.DataGrid.ColumnLabel(intGridCol, 1) = varData(0, intDataCol)
                intOffset = intOffset + 1
            End If
        Next
    Next
    'Fill in data values
    For intRowCounter = 1 To intPeriods
        For intColCounter = 1 To intUWYears
            intOffset = 1
            For intAcCounter = 1 To intAccounts
                If DisplayAccount(intAcCounter) = True Then
                    intGridCol = ((intColCounter - 1) * intNumVisAccounts) + intOffset
                    intDataCol = ((intColCounter - 1) * intAccounts) + intAcCounter
                    If IsNumeric(varData(intRowCounter, intDataCol)) = True Then
                        Call graControl.DataGrid.SetData _
                            (intRowCounter, intGridCol, _
                            CDbl(varData(intRowCounter, intDataCol)), False)
                        intOffset = intOffset + 1
                    Else
                        Call graControl.DataGrid.SetData _
                            (intRowCounter, intGridCol, _
                            0#, True)
                        intOffset = intOffset + 1
                    End If
                End If
            Next
        Next
    Next
Exit Sub
'Error Handler
FillChart_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub


