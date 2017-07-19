VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TriDataArrayOld"
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
'Public properties
'Private types
Private Type Account
    'Use acc as a prefix
    Name As String
    Display As Boolean
    Colour As Long
    Mask As String
End Type
Private Type Period
    'Use prd as a prefix
    Name As String
End Type
'Private properties
Private strColumnHeaders() As String
Private strRowHeaders() As String
Private curData() As Currency
Private accAccounts() As Account
Private intUWYears As Integer
Private intAccounts As Integer
Private prdPeriods() As Period
Private intDisplayPeriods As Integer
Private intYears As Integer
Private intPeriods As Integer
Private intReturn As Integer
Private strInstance As String
'Counter variables
Private intRowCounter As Integer
Private intColCounter As Integer
Private intAcCounter As Integer
Private intPeriodCounter As Integer
'Constants
Private Const conModule = "TriDataArray"
Private Const conAccounts = 9
Private Const conQuarters = 4
Private Const conPeriods = 12
Private Const conShowYears = 0
Private Const conShowQuarters = 1
Private Const conShowMonths = 2
'File constants
Private Const conSeperator = ","
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
Private Const errArray = 55001
Private Const errMismatch = 55002
Private Const errTooManyAccounts = 55003
Private Const errBadYearValue = 55004
Private Const errCannotHideLastAccount = 55005
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
'       a size equal to conAccounts
'   10. Total number of data columns is intUWYears * conAccounts (numbered from 1)
'   11. accAccounts() is used to determine
'       whether various display elements should be displayed FOR EACH UWYEAR
'   12. strRowHeaders() contains row header data in one dimension
'   13. strRowHeaders() starts from (1,0) at all times
'   14. Col(0) in strRowHeaders() contains Year headers
'   15. Col(1) in strRowHeaders() contains Period headers
'   16. Rows in strRowHeaders() are grouped into years with
'       a size equal to conPeriods
'   17. Total number of data rows is equal to intYears * intPeriods (numbered from 1)
'   18. prdPeriods() is used to determine
'       whether various display elements should be displayed FOR EACH YEAR
'
'-----------------------------------------------------------------------------

Private Sub Class_Initialize()
'Created by Robin G Brown, 6/5/97
'Default Behaviour
'Sub Declares
    'Error Trap
    On Error Resume Next
    'Initialise the arrays and default values
    strInstance = "Undefined"
    intUWYears = 15
    intAccounts = conAccounts
    intYears = 15
    intPeriods = conPeriods
    intDisplayPeriods = conShowYears
    Call InitialiseData
    Call InitialiseFormat
End Sub

Private Sub InitialiseData()
'Created by Robin G Brown, 7/5/97
'Initialise the data array
'Sub Declares
    'Error Trap
    On Error Resume Next
    'Column Headers
    ReDim strColumnHeaders(0 To 1, 1 To intYears * intAccounts)
    For intColCounter = 1 To intUWYears
        For intAcCounter = 1 To intAccounts
            strColumnHeaders(0, (intColCounter - 1) * intAccounts + (intAcCounter - 1) + 1) = ""
            strColumnHeaders(1, (intColCounter - 1) * intAccounts + (intAcCounter - 1) + 1) = ""
        Next
    Next
    'Row Headers
    ReDim strRowHeaders(1 To intYears * intPeriods, 0 To 1)
    For intRowCounter = 1 To intYears
        For intPeriodCounter = 1 To intPeriods
            strRowHeaders((intRowCounter - 1) * intPeriods + (intPeriodCounter - 1) + 1, 0) = ""
            strRowHeaders((intRowCounter - 1) * intPeriods + (intPeriodCounter - 1) + 1, 1) = ""
        Next
    Next
    'Currency data
    ReDim curData(1 To intYears, 1 To intUWYears * intAccounts)
    For intRowCounter = 1 To intYears
        For intColCounter = 1 To intUWYears * intAccounts
            curData(intRowCounter, intColCounter) = 0#
        Next
    Next
End Sub

Private Sub InitialiseFormat()
'Created by Robin G Brown, 7/5/97
'Initialise the displaycolumn array
'Sub Declares
    'Error Trap
    On Error Resume Next
    ReDim accAccounts(1 To intAccounts)
    For intAcCounter = 1 To intAccounts
        accAccounts(intAcCounter).Name = DefaultAccountName(intAcCounter)
        If InStr(accAccounts(intAcCounter).Name, "%") > 0 Then
            accAccounts(intAcCounter).Display = False
            accAccounts(intAcCounter).Mask = "0\%"
        Else
            accAccounts(intAcCounter).Display = True
            accAccounts(intAcCounter).Mask = "0"
        End If
        accAccounts(intAcCounter).Colour = vbBlack
    Next
    ReDim prdPeriods(1 To intPeriods)
    For intPeriodCounter = 1 To intPeriods
        prdPeriods(intPeriodCounter).Name = DefaultPeriodName(intPeriodCounter)
    Next
End Sub

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
    ReDim prdPeriods(1 To 1)
End Sub

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
        ThrowError errBadYearValue
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
        ThrowError errBadYearValue
    Else
        intYears = intNewValue
    End If
    Call InitialiseData
    Call InitialiseFormat
End Property

Property Get Periods() As Integer
'Created by Robin G Brown, 6/5/97
'Return the intPeriods value, depending on the state of intDisplayPeriods
'Function Declares
    'Error Trap
    On Error Resume Next
    Select Case intDisplayPeriods
    Case conShowYears
        Periods = 1
    Case conShowQuarters
        Periods = 4
    Case Else
        Periods = intPeriods
    End Select
End Property

Property Get DisplayPeriods() As Integer
'Created by Robin G Brown, 21/5/97
'Return the value of intDisplayPeriods
'Function Declares
    'Error Trap
    On Error Resume Next
    DisplayPeriods = intDisplayPeriods
End Property

Property Let DisplayPeriods(ByVal intDisplay As Integer)
'Created by Robin G Brown, 21/5/97
'Set the value of intDisplayPeriods
'Sub Declares
    'Error Trap
    On Error Resume Next
    intDisplayPeriods = intDisplay
End Property

Property Get Accounts() As Integer
'Created by Robin G Brown, 6/5/97
'Return the number of accounts
'Function Declares
    'Error Trap
    On Error Resume Next
    Accounts = intAccounts
End Property
'Property Let Accounts(ByVal intNewValue As Integer)
'DISABLED - you are not allowed to change the number of accounts - RGB/14/5/97
'End Property

Public Function NameOfAccount(ByVal intAccount As Integer) As String
'Created by Robin G Brown, 14/5/97
'Return the name of the account in strColumnheaders
'Sub Declares
    'Error Trap
    On Error Resume Next
    NameOfAccount = accAccounts(intAccount).Name
End Function

Public Sub AccountName(ByVal intAccount As Integer, ByVal strNewName As String)
'Created by Robin G Brown, 14/5/97
'Set the name of the account in strColumnheaders
'Sub Declares
    'Error Trap
    On Error Resume Next
    'Set the name
    accAccounts(intAccount).Name = strNewName
    'Fill in strColumnHeaders
    For intColCounter = 1 To intUWYears
        'strColumnHeaders(0, (intColCounter - 1) * intAccounts + (intAccount - 1) + 1) = ""
        strColumnHeaders(1, (intColCounter - 1) * intAccounts + (intAccount - 1) + 1) = strNewName
    Next
End Sub

Public Function VisibleAccounts() As Integer
'Created by Robin G Brown, 7/5/97
'Returns the number of visible accounts
'Function Declares
Dim intVisible As Integer
    'Error Trap
    On Error Resume Next
    intVisible = 0
    For intAcCounter = 1 To intAccounts
        If accAccounts(intAcCounter).Display = True Then
            intVisible = intVisible + 1
        End If
    Next
    VisibleAccounts = intVisible
End Function

Public Function LoadData(ByRef varArray() As Variant) As Boolean
'Created by Robin G Brown, 2/5/97
'Fill curData() from an array of data - varArray() _
    varArray must have lbounds of 1!
'Function Declares
Const conFunction = "LoadData"
Dim intRows As Integer
Dim intCols As Integer
    'Error Trap
    On Error GoTo LoadData_ErrorHandler
    'Read the size of the array
    intRows = UBound(varArray, 1)
    intCols = UBound(varArray, 1)
    'Now we know how big it is we can load the data into our own array
    ReDim curData(1 To intRows, 1 To intCols)
    For intRowCounter = 1 To intRows
        For intColCounter = 1 To intCols
            curData(intRowCounter, intColCounter) = _
                varArray(intRowCounter, intColCounter)
        Next
    Next
    LoadData = True
Exit Function
'Error Handler
LoadData_ErrorHandler:
    ReDim curData(1 To 1, 1 To 1)
    LoadData = False
    Exit Function
End Function

Public Sub CopyData(ByRef varArray() As Variant)
'Created by Robin G Brown, 2/5/97
'Copy curData into varArray, resizing varArray in the process
'Function Declares
Const conSub = "CopyData"
Dim intRows As Integer
Dim intCols As Integer
    'Error Trap
    On Error GoTo CopyData_ErrorHandler
    'Read the size of the array
    intRows = UBound(curData, 1)
    intCols = UBound(curData, 2)
    'Now we know how big it is we can load the data into varArray
    ReDim varArray(1 To intRows, 1 To intCols)
    For intRowCounter = 1 To intRows
        For intColCounter = 1 To intCols
            varArray(intRowCounter, intColCounter) = _
                curData(intRowCounter, intColCounter)
        Next
    Next
Exit Sub
'Error Handler
CopyData_ErrorHandler:
    ReDim varArray(1 To 1, 1 To 1)
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
        For intAcCounter = 1 To intAccounts
            If accAccounts(intAcCounter).Display = True Then
                intVisibleAccounts = intVisibleAccounts + 1
            End If
        Next
        If intVisibleAccounts > 1 Then
            accAccounts(intAccount).Display = False
        Else
            ThrowError errCannotHideLastAccount
        End If
    End If
End Function

Public Function ShowAllAccounts()
'Created by Robin G Brown, 7/5/97
'Sets accAccounts to show all account columns
'Sub Declares
    'Error Trap
    On Error Resume Next
    For intAcCounter = 1 To intAccounts
        accAccounts(intAcCounter).Display = True
    Next
End Function

Public Function GetAccountFromColumn(ByVal intColumn As Integer) As Integer
'Created by Robin G Brown, 13/5/97
'Return which account is in a given column, _
    assuming there are 2 fixed columns
'Function Declares
Dim intAccount As Integer
Dim intNumVisAccounts As Integer
Dim intVisAccounts As Integer
    'Error Trap
    On Error Resume Next
    'Adjust for 2 fixed columns
    intColumn = intColumn - 1
    intNumVisAccounts = VisibleAccounts
    intAccount = (intColumn Mod intNumVisAccounts)
    If intAccount = 0 Then
        intAccount = intNumVisAccounts
    End If
    'Go through all accounts counting visible ones, until we find the right one
    intVisAccounts = 0
    For intAcCounter = 1 To intAccounts
        If accAccounts(intAcCounter).Display = True Then
            intVisAccounts = intVisAccounts + 1
        End If
        If intVisAccounts = intAccount Then
            GetAccountFromColumn = intAcCounter
            Exit Function
        End If
    Next
    GetAccountFromColumn = 0
End Function

Private Function GetDataColumnFromColumn(ByVal intColumn As Integer) As Integer
'Created by Robin G Brown, 13/5/97
'Return which data column is used for a given grid column, _
    assuming that there are _2_ fixed columns!
'Function Declares
Dim intAccount As Integer
Dim intMultiplier As Integer
Dim intNumVisAccounts As Integer
Dim intVisAccounts As Integer
    'Error Trap
    On Error Resume Next
    'Adjust for the 2 fixed columns
    intColumn = intColumn - 1
    intNumVisAccounts = VisibleAccounts
    intAccount = intColumn Mod intNumVisAccounts
    If intAccount = 0 Then
        intAccount = intNumVisAccounts
    End If
    'Go through all accounts counting visible ones, until we find the right one
    intVisAccounts = 0
    For intAcCounter = 1 To intAccounts
        If accAccounts(intAcCounter).Display = True Then
            intVisAccounts = intVisAccounts + 1
        End If
        If intVisAccounts = intAccount Then
            Exit For
        End If
    Next
    intMultiplier = intColumn \ intNumVisAccounts
    If intColumn Mod intNumVisAccounts = 0 Then
        intMultiplier = intMultiplier - 1
    End If
    GetDataColumnFromColumn = intMultiplier * intAccounts + intAcCounter
End Function

Private Function GetUWYearFromColumn(ByVal intColumn As Integer) As Integer
'Created by Robin G Brown, 13/5/97
'Return which data column is used for a given column, _
    assuming there are 2 fixed columns
'Function Declares
Dim intAccount As Integer
Dim intMultiplier As Integer
Dim intNumVisAccounts As Integer
Dim intVisAccounts As Integer
    'Error Trap
    On Error Resume Next
    'Adjust for the 2 fixed columns
    intColumn = intColumn - 1
    intNumVisAccounts = VisibleAccounts
    intMultiplier = intColumn \ intNumVisAccounts
    If intColumn Mod intNumVisAccounts = 0 Then
        intMultiplier = intMultiplier - 1
    End If
    GetUWYearFromColumn = intMultiplier + 1
End Function

Public Function AccountColour(ByVal intAccount As Integer) As Long
'Created by Robin G Brown, 6/5/97
'Return the display colour of an account
'Function Declares
    'Error Trap
    On Error Resume Next
    AccountColour = accAccounts(intAccount).Colour
End Function

Public Sub SetAccountColour(ByVal intAccount As Integer, ByVal lngNewColour As Long)
'Created by Robin G Brown, 6/5/97
'Set the display colour of an account
'Function Declares
    'Error Trap
    On Error Resume Next
    accAccounts(intAccount).Colour = lngNewColour
End Sub

Public Function AccountMask(ByVal intAccount As Integer) As String
'Created by Robin G Brown, 6/5/97
'Return the display colour of an account
'Function Declares
    'Error Trap
    On Error Resume Next
    AccountMask = accAccounts(intAccount).Mask
End Function

Public Sub SetAccountMask(ByVal intAccount As Integer, ByVal strNewMask As String)
'Created by Robin G Brown, 6/5/97
'Set the mask of an account
'Function Declares
    'Error Trap
    On Error Resume Next
    accAccounts(intAccount).Mask = strNewMask
End Sub

Public Function DisplayColumn(ByVal intColumn As Integer) As Boolean
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
    intAccount = (intColumn Mod intAccounts)
    If intAccount = 0 Then
        intAccount = intAccounts
    End If
    DisplayColumn = accAccounts(intAccount).Display
End Function

Public Function DisplayAccount(ByVal intAccount As Integer) As Boolean
'Created by Robin G Brown, 6/5/97
'Return wether or not to display an account
'Function Declares
    'Error Trap
    On Error Resume Next
    DisplayAccount = accAccounts(intAccount).Display
End Function

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

Private Sub ThrowError(ByVal errNumber As Long)
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
    Case errCannotHideLastAccount
        Err.Raise errCannotHideLastAccount, conModule, "Cannot hide all accounts"
    Case Else
        Err.Raise 55000, conModule, "Undefined Error"
    End Select
End Sub

Public Sub CreateData(ByVal intStartYear As Integer)
'This sub should be private for 'release'
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
    ReDim strColumnHeaders(0 To 1, 1 To intUWYears * intAccounts)
    For intColCounter = 1 To intUWYears
        For intAcCounter = 1 To intAccounts
            strColumnHeaders(0, (intColCounter - 1) * intAccounts + (intAcCounter - 1) + 1) = CStr(intStartYear + intColCounter - 1)
            strColumnHeaders(1, (intColCounter - 1) * intAccounts + (intAcCounter - 1) + 1) = DefaultAccountName(intAcCounter)
        Next
    Next
    'Row Headers
    ReDim strRowHeaders(1 To intYears * intPeriods, 0 To 1)
    For intRowCounter = 1 To intYears
        For intPeriodCounter = 1 To intPeriods
            strRowHeaders((intRowCounter - 1) * intPeriods + (intPeriodCounter - 1) + 1, 0) = CStr(intRowCounter)
            strRowHeaders((intRowCounter - 1) * intPeriods + (intPeriodCounter - 1) + 1, 1) = DefaultPeriodName(intPeriodCounter)
        Next
    Next
    'Currency data
    ReDim curData(1 To intYears * intPeriods, 1 To intUWYears * intAccounts)
    For intColCounter = 1 To intUWYears
        For intAcCounter = 1 To intAccounts
            sngRandom = Rnd * 10
            For intRowCounter = 1 To intYears
                For intPeriodCounter = 1 To intPeriods
                    If intRowCounter <= (intYears + 1 - intColCounter) Then
                        If intRowCounter = 1 And intPeriodCounter = 1 Then
                            sngValue = sngRandom
                            sngPrevious = sngValue
                            sngDivisor = -1
                        Else
                            If sngDivisor < 1 Then sngDivisor = 0.5
                            sngValue = sngPrevious + (sngRandom / sngDivisor)
                            sngPrevious = sngValue
                            sngDivisor = sngDivisor + 1
                        End If
                        curData((intRowCounter - 1) * intPeriods + (intPeriodCounter - 1) + 1, (intColCounter - 1) * intAccounts + (intAcCounter - 1) + 1) _
                            = sngValue
                    Else
                        curData((intRowCounter - 1) * intPeriods + (intPeriodCounter - 1) + 1, (intColCounter - 1) * intAccounts + (intAcCounter - 1) + 1) _
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

Private Function SumYear(ByVal intYear As Integer, ByVal intDataColumn As Integer) As Currency
'Created by Robin G Brown, 21/5/97
'Returns the sum of the periods for the given intYear and intDataColumn
'Function Declares
Dim intSumCounter As Integer
Dim curSum As Currency
    'Error Trap
    On Error GoTo SumYear_ErrorHandler
    curSum = 0#
    For intSumCounter = ((intYear - 1) * intPeriods) + 1 To intYear * intPeriods
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
'Returns the sum of the periods for the given intYear and intDataColumn
'Function Declares
Dim intSumCounter As Integer
Dim curSum As Currency
Dim intStartMonth As Integer
Dim intEndMonth As Integer
    'Error Trap
    On Error GoTo SumQuarter_ErrorHandler
    curSum = 0#
    intStartMonth = ((intYear - 1) * intPeriods) + 1
    intEndMonth = intYear * intPeriods
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

Public Sub FillGrid(ByRef grdControl As vsFlexArray) 'MSFlexGrid)
'Created by Robin G Brown, 2/5/97
'Fill the given grid with data from the arrays _
    colouring, etc. as appropriate
'Sub Declares
Const conSub = "FillGrid"
Dim intGridCol As Integer
Dim intDataCol As Integer
Dim intGridRow As Integer
Dim intDataRow As Integer
Dim intNumVisAccounts As Integer
Dim intOffset As Integer
Dim curCurrentData As Currency
    'Error Trap
    On Error GoTo FillGrid_ErrorHandler
    Screen.MousePointer = vbHourglass
    With grdControl
        'Set up grdControl
        .FillStyle = flexFillRepeat
        '.Clear
        'Allow rows and cols only to merge
        .MergeCells = flexMergeRestrictAll
        'Set only the top row and first col to merge
        .MergeRow(0) = True
        .MergeCol(0) = True
        'Set the number of fixed cols and rows
        .FixedCols = 0
        .FixedRows = 0
        'Set the number of cols and rows
        intNumVisAccounts = VisibleAccounts
        .Cols = (intUWYears * intNumVisAccounts) + 2
        'Temporarily set a value for rows
        .Rows = 2
        'Write the col headers
        For intColCounter = 1 To intUWYears
            intOffset = 2
            For intAcCounter = 1 To intAccounts
                If DisplayAccount(intAcCounter) = True Then
                    intGridCol = ((intColCounter - 1) * intNumVisAccounts) + intOffset
                    intDataCol = ((intColCounter - 1) * intAccounts) + intAcCounter
                    .TextMatrix(0, intGridCol) = strColumnHeaders(0, intDataCol)
                    .TextMatrix(1, intGridCol) = strColumnHeaders(1, intDataCol)
                    intOffset = intOffset + 1
                End If
            Next
        Next
        'Now fill in all the rows, col by col
        If intDisplayPeriods = conShowYears Then
            'Display yearly values
            .Rows = intYears + 2
            'Write the row headers
            For intRowCounter = 1 To intYears
                .TextMatrix(intRowCounter + 1, 0) = strRowHeaders(((intRowCounter - 1) * intPeriods) + 1, 0)
                .TextMatrix(intRowCounter + 1, 1) = "Year"
            Next
            'Set the number of fixed cols and rows
            .FixedCols = 2
            .FixedRows = 2
            'Now fill in the data, column by column
            For intColCounter = 1 To intUWYears
                intOffset = 2
                For intAcCounter = 1 To intAccounts
                    If DisplayAccount(intAcCounter) = True Then
                        'Work out where the data comes from and where it goes in the grid
                        intGridCol = ((intColCounter - 1) * intNumVisAccounts) + intOffset
                        intDataCol = ((intColCounter - 1) * intAccounts) + intAcCounter
                        'Write the data, row by row
                        For intRowCounter = 1 To intYears
                            curCurrentData = SumYear(intRowCounter, intDataCol)
                            'Write the data to the grid, applying any format required
                            If curCurrentData <> 0# Then
                                .TextMatrix(intRowCounter + 1, intGridCol) = _
                                    Format(curCurrentData, accAccounts(intAcCounter).Mask)
                                'Format the cells
                                .Col = intGridCol
                                .Row = intRowCounter + 1 'grdControl.FixedRows
                                '.RowSel = grdControl.Rows - 1
                                .CellForeColor = accAccounts(intAcCounter).Colour
                                If intColCounter Mod 2 = 0 Then
                                    .CellBackColor = &HE0E0E0
                                Else
                                    .CellBackColor = vbWhite
                                End If
                            Else
                                'Show a blank if the value is 0
                                .TextMatrix(intRowCounter + 1, intGridCol) = ""
                            End If
                        Next
                        'Set the offset for the next col...
                        intOffset = intOffset + 1
                    End If
                Next
            Next
        ElseIf intDisplayPeriods = conShowQuarters Then
            'Display quarterly values
            .Rows = (intYears * conQuarters) + 2
            'Write the row headers
            For intRowCounter = 1 To intYears
                For intPeriodCounter = 1 To conQuarters
                    intGridRow = ((intRowCounter - 1) * conQuarters) + intPeriodCounter
                    intDataRow = ((intRowCounter - 1) * intPeriods) + intPeriodCounter
                    .TextMatrix(intGridRow + 1, 0) = strRowHeaders(intDataRow, 0)
                    .TextMatrix(intGridRow + 1, 1) = "Q" & intPeriodCounter
                Next
            Next
            'Set the number of fixed cols and rows
            .FixedCols = 2
            .FixedRows = 2
            'Now fill in the data, column by column
            For intColCounter = 1 To intUWYears
                intOffset = 2
                For intAcCounter = 1 To intAccounts
                    If DisplayAccount(intAcCounter) = True Then
                        'Work out where the data comes from and where it goes in the grid
                        intGridCol = ((intColCounter - 1) * intNumVisAccounts) + intOffset
                        intDataCol = ((intColCounter - 1) * intAccounts) + intAcCounter
                        'Write the data, row by row
                        For intRowCounter = 1 To intYears
                            For intPeriodCounter = 1 To conQuarters
                                intGridRow = ((intRowCounter - 1) * conQuarters) + intPeriodCounter
                                curCurrentData = SumQuarter(intRowCounter, intPeriodCounter, intDataCol)
                                'Write the data to the grid, applying any format required
                                If curCurrentData <> 0# Then
                                    .TextMatrix(intGridRow + 1, intGridCol) = _
                                        Format(curCurrentData, accAccounts(intAcCounter).Mask)
                                    'Format the cells
                                    .Col = intGridCol
                                    .Row = intGridRow + 1
                                    .CellForeColor = accAccounts(intAcCounter).Colour
                                    If intColCounter Mod 2 = 0 Then
                                        .CellBackColor = &HE0E0E0
                                    Else
                                        .CellBackColor = vbWhite
                                    End If
                                Else
                                    'Show a blank if the value is 0
                                    .TextMatrix(intGridRow + 1, intGridCol) = ""
                                End If
                            Next
                        Next
                        'Set the offset for the next col...
                        intOffset = intOffset + 1
                    End If
                Next
            Next
        Else
            'Display all months
            .Rows = (intYears * intPeriods) + 2
            'Write the row headers
            For intRowCounter = 1 To intYears
                For intPeriodCounter = 1 To intPeriods
                    intDataRow = ((intRowCounter - 1) * intPeriods) + intPeriodCounter
                    .TextMatrix(intDataRow + 1, 0) = strRowHeaders(intDataRow, 0)
                    .TextMatrix(intDataRow + 1, 1) = strRowHeaders(intDataRow, 1)
                Next
            Next
            'Set the number of fixed cols and rows
            .FixedCols = 2
            .FixedRows = 2
            'Now fill in the data, column by column
            For intColCounter = 1 To intUWYears
                intOffset = 2
                For intAcCounter = 1 To intAccounts
                    If DisplayAccount(intAcCounter) = True Then
                        'Work out where the data comes from and where it goes in the grid
                        intGridCol = ((intColCounter - 1) * intNumVisAccounts) + intOffset
                        intDataCol = ((intColCounter - 1) * intAccounts) + intAcCounter
                        'Write the data, row by row
                        For intRowCounter = 1 To intYears
                            For intPeriodCounter = 1 To intPeriods
                                intDataRow = ((intRowCounter - 1) * intPeriods) + intPeriodCounter
                                'Write the data to the grid, applying any format required
                                If curData(intDataRow, intDataCol) <> 0# Then
                                    .TextMatrix(intDataRow + 1, intGridCol) = _
                                        Format(curData(intDataRow, intDataCol), accAccounts(intAcCounter).Mask)
                                    'Format the cells
                                    .Col = intGridCol
                                    .Row = intDataRow + 1
                                    .CellForeColor = accAccounts(intAcCounter).Colour
                                    If intColCounter Mod 2 = 0 Then
                                        .CellBackColor = &HE0E0E0
                                    Else
                                        .CellBackColor = vbWhite
                                    End If
                                Else
                                    'Show a blank if the value is 0
                                    .TextMatrix(intDataRow + 1, intGridCol) = ""
                                End If
                            Next
                        Next
                        'Set the offset for the next col...
                        intOffset = intOffset + 1
                    End If
                Next
            Next
        End If
    End With
    Screen.MousePointer = vbDefault
Exit Sub
'Error Handler
FillGrid_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Public Sub FillChart(ByRef graGraphControl As MSChart, ByRef grdGridControl As MSFlexGrid, ByVal intAccount As Integer, ByVal intColumn As Integer)
'Created by Robin G Brown, 6/5/97
'This function copied from the Chart Control Library
'Set up a chart from the curData() array, _
    if intColumn=0 then chart all columns representing Account(intAccount), _
    else chart only intColumn, _
    assuming there are 2 fixed columns but no need to adjust for them, _
    as this is done in other routines
'Sub Declares
Const conSub = "FillChart"
Dim intCols As Integer
Dim intRows As Integer
Dim intChartCol As Integer
Dim intGridCol As Integer
Dim intDataRow As Integer
Dim intOffset As Integer
Dim intRed As Integer
Dim intGreen As Integer
Dim intBlue As Integer
Dim curCurrentData As Currency
Dim intNumVisAccounts As Integer
Const conNumVisAccounts = 1
    'Error Trap
    On Error GoTo FillChart_ErrorHandler
    With graGraphControl
        'Set up the chart first
        .chartType = VtChChartType2dLine
        'Decide what to show
        If intColumn = 0 Then
            'This will show multiple columns of data
            'Set the size of the datagrid first
            intRows = grdGridControl.Rows - 2
            intCols = intUWYears
            Call .DataGrid.SetSize(intRows, intCols, intRows, intCols)
            'Fill in labels for rows
            For intRowCounter = 1 To intRows
                .DataGrid.RowLabel(intRowCounter, 1) = _
                    grdGridControl.TextMatrix(intRowCounter + 1, 0)
                .DataGrid.RowLabel(intRowCounter, 2) = _
                    grdGridControl.TextMatrix(intRowCounter + 1, 1)
            Next
            'Fill in labels for columns
            intNumVisAccounts = VisibleAccounts
            For intColCounter = 1 To intUWYears
                intOffset = 1
                For intAcCounter = 1 To intAccounts
                    If intAcCounter = intAccount Then
                        intChartCol = ((intColCounter - 1) * conNumVisAccounts) + intOffset
                        intGridCol = ((intColCounter - 1) * intNumVisAccounts) + intAcCounter + 1
                        .DataGrid.ColumnLabel(intChartCol, 1) = _
                            grdGridControl.TextMatrix(0, intGridCol) _
                            & "," & grdGridControl.TextMatrix(1, intGridCol)
                        intOffset = intOffset + 1
                    End If
                Next
            Next
            'Fill in data values
            For intRowCounter = 1 To intRows
                For intColCounter = 1 To intUWYears
                    intChartCol = ((intColCounter - 1) * conNumVisAccounts) + 1
                    intGridCol = ((intColCounter - 1) * intNumVisAccounts) + intAccount + 1
                    If grdGridControl.TextMatrix(intRowCounter + 1, intGridCol) <> "" Then
                        Call .DataGrid.SetData _
                            (intRowCounter, intChartCol, _
                            CDbl(grdGridControl.TextMatrix(intRowCounter + 1, intGridCol)), False)
                        'intOffset = intOffset + 1
                    Else
                        Call .DataGrid.SetData(intRowCounter, intChartCol, 0#, True)
                        'intOffset = intOffset + 1
                    End If
                    'Set up some colours
                    intRed = RedFromRGB(accAccounts(intAccount).Colour)
                    intGreen = GreenFromRGB(accAccounts(intAccount).Colour)
                    intBlue = BlueFromRGB(accAccounts(intAccount).Colour)
                    'Set the colour
                    With .Plot.SeriesCollection(intChartCol).Pen
                        .VtColor.Set intRed, intGreen, intBlue
                        .Width = 0
                    End With
                Next
            Next
        Else
            'This will show a single column of data
            intRows = grdGridControl.Rows - 2
            intCols = 1
            Call .DataGrid.SetSize(intRows, intCols, intRows, intCols)
            'Fill in labels for rows
            For intRowCounter = 1 To intRows
                .DataGrid.RowLabel(intRowCounter, 1) = _
                    grdGridControl.TextMatrix(intRowCounter + 1, 0)
                .DataGrid.RowLabel(intRowCounter, 2) = _
                    grdGridControl.TextMatrix(intRowCounter + 1, 1)
            Next
            'Fill in labels for column
            .DataGrid.ColumnLabel(1, 1) = _
                grdGridControl.TextMatrix(0, intColumn) _
                & "," & grdGridControl.TextMatrix(1, intColumn)
            'Fill in data values
            For intRowCounter = 1 To intRows
                If grdGridControl.TextMatrix(intRowCounter + 1, intColumn) <> "" Then
                    Call .DataGrid.SetData _
                        (intRowCounter, 1, _
                        CDbl(grdGridControl.TextMatrix(intRowCounter + 1, intColumn)), False)
                Else
                    Call .DataGrid.SetData(intRowCounter, 1, 0#, True)
                End If
                'Set up some colours
                intRed = RedFromRGB(accAccounts(GetAccountFromColumn(intColumn)).Colour)
                intGreen = GreenFromRGB(accAccounts(GetAccountFromColumn(intColumn)).Colour)
                intBlue = BlueFromRGB(accAccounts(GetAccountFromColumn(intColumn)).Colour)
                'Set the colour
                With .Plot.SeriesCollection(1).Pen
                    .VtColor.Set intRed, intGreen, intBlue
                    .Width = 0
                End With
            Next
        End If
    End With
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

Public Sub OldFillChart(ByRef graControl As MSChart, ByVal intAccount As Integer, ByVal intColumn As Integer)
'Created by Robin G Brown, 6/5/97
'This function copied from the Chart Control Library
'Set up a chart from the curData() array, _
    if intColumn=0 then chart all columns representing Account(intAccount), _
    else chart only intColumn
'Sub Declares
Const conSub = "FillChart"
Dim intCols As Integer
Dim intRows As Integer
Dim intChartCol As Integer
Dim intDataCol As Integer
Dim intDataRow As Integer
Dim intOffset As Integer
Dim intRed As Integer
Dim intGreen As Integer
Dim intBlue As Integer
Const conNumVisAccounts = 1
    'Error Trap
    On Error GoTo FillChart_ErrorHandler
    With graControl
        'Set up the chart first
        .chartType = VtChChartType2dLine
        'Set the size of the datagrid first
        If intColumn = 0 Then
            'This will show multiple columns of data
            intRows = intYears * intPeriods
            intCols = intUWYears
            Call .DataGrid.SetSize(intRows, intCols, intRows, intCols)
            'Fill in labels for rows
            For intRowCounter = 1 To intRows
                .DataGrid.RowLabel(intRowCounter, 1) = _
                    CStr(strRowHeaders(intRowCounter, 0))
                .DataGrid.RowLabel(intRowCounter, 2) = _
                    CStr(strRowHeaders(intRowCounter, 1))
            Next
            'Fill in labels for columns
            For intColCounter = 1 To intUWYears
                intOffset = 1
                For intAcCounter = 1 To intAccounts
                    If intAcCounter = intAccount Then
                        intChartCol = ((intColCounter - 1) * conNumVisAccounts) + intOffset
                        intDataCol = ((intColCounter - 1) * intAccounts) + intAcCounter
                        .DataGrid.ColumnLabel(intChartCol, 1) = strColumnHeaders(0, intDataCol) & "," & strColumnHeaders(1, intDataCol)
                        intOffset = intOffset + 1
                    End If
                Next
            Next
            'Fill in data values
            For intRowCounter = 1 To intRows
                For intColCounter = 1 To intUWYears
                    intOffset = 1
                    For intAcCounter = 1 To intAccounts
                        If intAcCounter = intAccount Then
                            intChartCol = ((intColCounter - 1) * conNumVisAccounts) + intOffset
                            intDataCol = ((intColCounter - 1) * intAccounts) + intAcCounter
                            If curData(intRowCounter, intDataCol) <> 0# Then
                                Call .DataGrid.SetData _
                                    (intRowCounter, intChartCol, _
                                    CDbl(curData(intRowCounter, intDataCol)), False)
                                intOffset = intOffset + 1
                            Else
                                Call .DataGrid.SetData _
                                    (intRowCounter, intChartCol, _
                                    0#, True)
                                intOffset = intOffset + 1
                            End If
                            'Set up some colours
                            intRed = RedFromRGB(accAccounts(intAcCounter).Colour)
                            intGreen = GreenFromRGB(accAccounts(intAcCounter).Colour)
                            intBlue = BlueFromRGB(accAccounts(intAcCounter).Colour)
                            'Set the colour
                            With .Plot.SeriesCollection(intChartCol).Pen
                                .VtColor.Set intRed, intGreen, intBlue
                                .Width = 0
                            End With
                        End If
                    Next
                Next
            Next
        Else
            'This will show a single column of data
            intCols = 1
            intDataCol = GetDataColumnFromColumn(intColumn)
            intRows = intYears * intPeriods
            Call .DataGrid.SetSize(intRows, intCols, intRows, intCols)
            'Fill in labels for rows
            For intRowCounter = 1 To intRows
                .DataGrid.RowLabel(intRowCounter, 1) = _
                    CStr(strRowHeaders(intRowCounter, 0))
                .DataGrid.RowLabel(intRowCounter, 2) = _
                    CStr(strRowHeaders(intRowCounter, 1))
            Next
            'Fill in labels for column
            .DataGrid.ColumnLabel(1, 1) = strColumnHeaders(0, intDataCol) & "," & strColumnHeaders(1, intDataCol)
            'Fill in data values
            For intRowCounter = 1 To intRows
                If curData(intRowCounter, intDataCol) <> 0# Then
                    Call .DataGrid.SetData _
                        (intRowCounter, 1, _
                        CDbl(curData(intRowCounter, intDataCol)), False)
                Else
                    Call .DataGrid.SetData _
                        (intRowCounter, 1, _
                        0#, True)
                End If
                'Set up some colours
                intRed = RedFromRGB(accAccounts(GetAccountFromColumn(intColumn)).Colour)
                intGreen = GreenFromRGB(accAccounts(GetAccountFromColumn(intColumn)).Colour)
                intBlue = BlueFromRGB(accAccounts(GetAccountFromColumn(intColumn)).Colour)
                'Set the colour
                With .Plot.SeriesCollection(1).Pen
                    .VtColor.Set intRed, intGreen, intBlue
                    .Width = 0
                End With
            Next
        End If
    End With
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

Public Sub FileSave(ByVal strFileAndPath As String)
'Created by Robin G Brown, 19/5/97
'Save the data and formatting to a file
'Sub Declares
Const conSub = "FileSave"
Dim intFileNumber As Integer
Dim intDataCol As Integer
Dim intDataRow As Integer
Dim strDataLine As String
Dim strSecondDataLine As String
    'Error Trap
    On Error GoTo FileSave_ErrorHandler
    'Check for a duplicate file first
    If CheckForDuplicate(strFileAndPath) = False Then
        Exit Sub
    End If
    'Next open the file
    intFileNumber = FreeFile()
    Open strFileAndPath For Output Access Write Lock Read Write As #intFileNumber
    'Save the header data
    'Column headers
    Print #intFileNumber, "COLUMNHEADER" & ""
    strDataLine = ""
    strSecondDataLine = ""
    For intColCounter = 1 To intUWYears
        For intAcCounter = 1 To intAccounts
            intDataCol = ((intColCounter - 1) * intAccounts) + intAcCounter
            strDataLine = strDataLine & strColumnHeaders(0, intDataCol) & conSeperator
            strSecondDataLine = strSecondDataLine & strColumnHeaders(1, intDataCol) & conSeperator
        Next
    Next
    strDataLine = Left$(strDataLine, Len(strDataLine) - 1) & ""
    strSecondDataLine = Left$(strSecondDataLine, Len(strSecondDataLine) - 1) & ""
    Print #intFileNumber, strDataLine
    Print #intFileNumber, strSecondDataLine
    'Row headers
    Print #intFileNumber, "ROWHEADER" & ""
    strDataLine = ""
    For intRowCounter = 1 To intYears
        For intPeriodCounter = 1 To intPeriods
            intDataRow = ((intRowCounter - 1) * intPeriods) + intPeriodCounter
            strDataLine = strDataLine & strRowHeaders(intDataRow, 0) & conSeperator
            strSecondDataLine = strSecondDataLine & strRowHeaders(intDataRow, 1) & conSeperator
        Next
    Next
    strDataLine = Left$(strDataLine, Len(strDataLine) - 1) & ""
    strSecondDataLine = Left$(strSecondDataLine, Len(strSecondDataLine) - 1) & ""
    Print #intFileNumber, strDataLine
    Print #intFileNumber, strSecondDataLine
    'Save the formatting data
    Print #intFileNumber, "FORMAT" & ""
    'Account names
    strDataLine = ""
    For intColCounter = 1 To UBound(accAccounts, 1)
        strDataLine = strDataLine & accAccounts(intColCounter).Name & conSeperator
    Next
    strDataLine = Left$(strDataLine, Len(strDataLine) - 1) & ""
    Print #intFileNumber, strDataLine
    'Display accounts
    strDataLine = ""
    For intColCounter = 1 To UBound(accAccounts, 1)
        strDataLine = strDataLine & accAccounts(intColCounter).Display & conSeperator
    Next
    strDataLine = Left$(strDataLine, Len(strDataLine) - 1) & ""
    Print #intFileNumber, strDataLine
    'Account colours
    strDataLine = ""
    For intColCounter = 1 To UBound(accAccounts, 1)
        strDataLine = strDataLine & accAccounts(intColCounter).Colour & conSeperator
    Next
    strDataLine = Left$(strDataLine, Len(strDataLine) - 1) & ""
    Print #intFileNumber, strDataLine
    'Account format masks
    strDataLine = ""
    For intColCounter = 1 To UBound(accAccounts, 1)
        strDataLine = strDataLine & accAccounts(intColCounter).Mask & conSeperator
    Next
    strDataLine = Left$(strDataLine, Len(strDataLine) - 1) & ""
    Print #intFileNumber, strDataLine
    'Save the actual data
    Print #intFileNumber, "DATA" & ""
    For intRowCounter = 1 To intYears
        strDataLine = ""
        For intColCounter = 1 To intUWYears
            For intAcCounter = 1 To UBound(accAccounts, 1)
                intDataCol = ((intColCounter - 1) * intAccounts) + intAcCounter
                strDataLine = strDataLine & curData(intRowCounter, intDataCol) & conSeperator
            Next
        Next
        strDataLine = Left$(strDataLine, Len(strDataLine) - 1) & ""
        Print #intFileNumber, strDataLine
    Next
    'Close the file
    Close #intFileNumber
Exit Sub
'Error Handler
FileSave_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Close #intFileNumber
    Exit Sub
End Sub

Public Sub FileOpen(ByVal strFileAndPath As String)
'Created by Robin G Brown, 19/5/97
'Retrieve the formatting and data from a file
'Sub Declares
Const conSub = "FileOpen"
Dim intFileNumber As Integer
Dim intLineType As Integer
Dim strDataLine As String
Dim intDataLineCounter As Integer
Dim strLineArray() As String
    'Error Trap
    'On Error GoTo FileOpen_ErrorHandler
    'Open the file
    intFileNumber = FreeFile()
    Open strFileAndPath For Input Access Read Lock Read Write As #intFileNumber
    'Read the first line
    intDataLineCounter = 0
    Line Input #intFileNumber, strDataLine
    Do While Not EOF(intFileNumber)
        'Act on the contents of that line
        intLineType = GetLineInputType(strDataLine)
        Select Case intLineType
        Case conColHead1Line
            'Retrieve the header data
            ReDim strLineArray(1 To CharactersInString(strDataLine, conSeperator) + 1)
            intReturn = WriteCSVToArray(strDataLine, strLineArray())
            For intColCounter = 1 To UBound(strLineArray, 1)
                strColumnHeaders(0, intColCounter) = strLineArray(intColCounter)
            Next
        Case conColHead2Line
            'Retrieve the header data
            ReDim strLineArray(1 To CharactersInString(strDataLine, conSeperator) + 1)
            intReturn = WriteCSVToArray(strDataLine, strLineArray())
            For intColCounter = 1 To UBound(strLineArray, 1)
                strColumnHeaders(1, intColCounter) = strLineArray(intColCounter)
            Next
        Case conRowHead1Line
            'Retrieve the header data
            ReDim strLineArray(1 To CharactersInString(strDataLine, conSeperator) + 1)
            intReturn = WriteCSVToArray(strDataLine, strLineArray())
            For intRowCounter = 1 To UBound(strLineArray, 1)
                strRowHeaders(intRowCounter, 0) = strLineArray(intRowCounter)
            Next
        Case conRowHead2Line
            'Retrieve the header data
            ReDim strLineArray(1 To CharactersInString(strDataLine, conSeperator) + 1)
            intReturn = WriteCSVToArray(strDataLine, strLineArray())
            For intRowCounter = 1 To UBound(strLineArray, 1)
                strRowHeaders(intRowCounter, 1) = strLineArray(intRowCounter)
            Next
        Case conAccNameLine
            'Retrieve the formatting data
            ReDim strLineArray(1 To CharactersInString(strDataLine, conSeperator) + 1)
            intReturn = WriteCSVToArray(strDataLine, strLineArray())
            For intColCounter = 1 To UBound(strLineArray, 1)
                accAccounts(intColCounter).Name = strLineArray(intColCounter)
            Next
        Case conAccDisplayLine
            'Retrieve the formatting data
            ReDim strLineArray(1 To CharactersInString(strDataLine, conSeperator) + 1)
            intReturn = WriteCSVToArray(strDataLine, strLineArray())
            For intColCounter = 1 To UBound(strLineArray, 1)
                accAccounts(intColCounter).Display = strLineArray(intColCounter)
            Next
        Case conAccColourLine
            'Retrieve the formatting data
            ReDim strLineArray(1 To CharactersInString(strDataLine, conSeperator) + 1)
            intReturn = WriteCSVToArray(strDataLine, strLineArray())
            For intColCounter = 1 To UBound(strLineArray, 1)
                accAccounts(intColCounter).Colour = strLineArray(intColCounter)
            Next
        Case conAccFormatLine
            'Retrieve the formatting data
            ReDim strLineArray(1 To CharactersInString(strDataLine, conSeperator) + 1)
            intReturn = WriteCSVToArray(strDataLine, strLineArray())
            For intColCounter = 1 To UBound(strLineArray, 1)
                accAccounts(intColCounter).Mask = strLineArray(intColCounter)
            Next
        Case conDataLine
            'Retrieve the actual data
            ReDim strLineArray(1 To CharactersInString(strDataLine, conSeperator) + 1)
            intDataLineCounter = intDataLineCounter + 1
            intReturn = WriteCSVToArray(strDataLine, strLineArray())
            For intColCounter = 1 To UBound(strLineArray, 1)
                curData(intDataLineCounter, intColCounter) = CCur(strLineArray(intColCounter))
            Next
        Case Else
            'Ignore this line
        End Select
        'Read the next line
        Line Input #intFileNumber, strDataLine
    Loop
    'Close the file
    Close #intFileNumber
Exit Sub
'Error Handler
FileOpen_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Close #intFileNumber
    Exit Sub
End Sub

Private Function GetLineInputType(ByVal strDataLine As String) As Integer
'Created by Robin G Brown, 20/5/97
'Given a string, returns various answers depending on what the previous answer was
'Function Declares
Const conFunction = "GetLineInputType"
Static intPreviousLineType As Integer
    'Error Trap
    On Error GoTo GetLineInputType_ErrorHandler
    If strDataLine = "" Then
        'Ignore blank lines
        Exit Function
    End If
    Select Case intPreviousLineType
    Case conUnknownLine
        'The last line was unknow so keep going _
            until we find a line that is recognised
        If InStr(1, strDataLine, "COLUMNHEADER") > 0 Then
            intPreviousLineType = conColHeaderLine
            GetLineInputType = conColHeaderLine
        ElseIf InStr(1, strDataLine, "ROWHEADER") > 0 Then
            intPreviousLineType = conRowHeaderLine
            GetLineInputType = conRowHeaderLine
        ElseIf InStr(1, strDataLine, "FORMAT") > 0 Then
            intPreviousLineType = conFormatHeaderLine
            GetLineInputType = conFormatHeaderLine
        ElseIf InStr(1, strDataLine, "DATA") > 0 Then
            intPreviousLineType = conDataHeaderLine
            GetLineInputType = conDataHeaderLine
        Else
            intPreviousLineType = conUnknownLine
            GetLineInputType = conUnknownLine
        End If
    Case 0
        If InStr(1, strDataLine, "COLUMNHEADER") > 0 Then
            intPreviousLineType = conColHeaderLine
            GetLineInputType = conColHeaderLine
        End If
    Case conColHeaderLine
        intPreviousLineType = conColHead1Line
        GetLineInputType = conColHead1Line
    Case conColHead1Line
        intPreviousLineType = conColHead2Line
        GetLineInputType = conColHead2Line
    Case conColHead2Line
        intPreviousLineType = conRowHeaderLine
        GetLineInputType = conRowHeaderLine
    Case conRowHeaderLine
        intPreviousLineType = conRowHead1Line
        GetLineInputType = conRowHead1Line
    Case conRowHead1Line
        intPreviousLineType = conRowHead2Line
        GetLineInputType = conRowHead2Line
    Case conRowHead2Line
        intPreviousLineType = conFormatHeaderLine
        GetLineInputType = conFormatHeaderLine
    Case conFormatHeaderLine
        intPreviousLineType = conAccNameLine
        GetLineInputType = conAccNameLine
    Case conAccNameLine
        intPreviousLineType = conAccDisplayLine
        GetLineInputType = conAccDisplayLine
    Case conAccDisplayLine
        intPreviousLineType = conAccColourLine
        GetLineInputType = conAccColourLine
    Case conAccColourLine
        intPreviousLineType = conAccFormatLine
        GetLineInputType = conAccFormatLine
    Case conAccFormatLine
        intPreviousLineType = conDataHeaderLine
        GetLineInputType = conDataHeaderLine
    Case conDataHeaderLine
        intPreviousLineType = conDataLine
        GetLineInputType = conDataLine
    Case conDataHeaderLine
        intPreviousLineType = conDataLine
        GetLineInputType = conDataLine
    Case Else
        intPreviousLineType = conUnknownLine
        GetLineInputType = conUnknownLine
    End Select
Exit Function
'Error Handler
GetLineInputType_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        'MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conFunction
    End Select
    GetLineInputType = conUnknownLine
    Exit Function
End Function



