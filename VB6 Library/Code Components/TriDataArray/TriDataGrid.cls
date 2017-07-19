VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TriDataGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Created by Robin G Brown, 18th September 1997
'-----------------------------------------------------------------------------
'   This class module contains code for _
    holding just the data for a triangulation report
'   Use tdg as a prefix for this class
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conClass = "TriDataGrid"
'Private types
Private Type Account
    'Use acc as a prefix
    Name As String
    Display As Boolean
    Mask As String
End Type
'Private properties
Private lngCounter As Long
Private lngReturn As Long
Private strColumnHeaders() As String
Private strRowHeaders() As String
Private curData() As Currency
Private accAccounts() As Account
Private intUWYears As Integer
Private intYears As Integer
'Counter variables
Private lngRowCounter As Long
Private lngColCounter As Long
Private intAcCounter As Integer
Private intMonthcounter As Integer
'Constants
Private Const conAccounts = 9
Private Const conQuarters = 4
Private Const conMonths = 12
Private Const conNumFixedCols = 3
Private Const conNumFixedRows = 2
'File constants
Private Const conSeperator = ","
Private Const conColon = ":"

Private Sub Class_Initialize()
'Created by Robin G Brown, 6/5/97
'Default Behaviour
'Sub Declares
    'Error Trap
    On Error Resume Next
    'Initialise the arrays and default values
    intUWYears = 15
    intYears = 15
    Call InitialiseData
    Call InitialiseFormat
End Sub

'-----------------------------------------------------------------------------
'   Event Handlers
'-----------------------------------------------------------------------------

'-----------------------------------------------------------------------------
'   Public Functions and Methods
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

'-----------------------------------------------------------------------------
'   Properties
'-----------------------------------------------------------------------------

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
End Sub




