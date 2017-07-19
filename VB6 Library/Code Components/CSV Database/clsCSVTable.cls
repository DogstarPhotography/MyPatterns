VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSVTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Created by Robin G Brown, 5th September 1997
'-----------------------------------------------------------------------------
'   This class module contains code for _
    a CSV records table
'   Use ctb as a prefix for this class
'-----------------------------------------------------------------------------
'   strDataLines is a 2 dimensional array for storing data read from a file
'   strDataLines(col, row)
'   Row 0 of strDataLines can be used to store field names
'   lngCurrentLine is a 'pointer' to the current record
'   lngSortOrder is an array listing the order of the fields to sort in
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Type CLASSFLAGS
    'Use flg as a prefix
    FieldNames As Boolean
    RecordFound As Boolean
End Type
Private Const conClass = "CSVTable"
Private Const conCSVTableSeparator = ","
Private Const conFindFirst = 1
Private Const conFindNext = 2
Private Const conFindLast = 3
Private Const conFindPrevious = 4
'Private properties
Private lngFieldCounter As Long
Private lngReturn As Long
Private strFileName As String
Private strDataLines() As String
Private lngSortOrder() As Long
Private lngCurrentLine As Long
Private lngRecordCount As Long
Private lngBookmark As Long
Private flgTableFlags As CLASSFLAGS

Private Sub Class_Initialize()
'Created by Robin G Brown, 5/9/97
'Set all the arrays and variables to the initial 'empty' state
'Sub Declares
    'Error Trap
    On Error Resume Next
    ReDim strDataLines(0 To 0, 0 To 0)
    ReDim strDataLines(0 To 0)
    strFileName = ""
    lngRecordCount = 0
    lngCurrentLine = -1
    With flgTableFlags
        .FieldNames = False
        .RecordFound = False
    End With
End Sub

'-----------------------------------------------------------------------------
'   Public Functions and Methods
'-----------------------------------------------------------------------------

Public Sub FileOpen(ByVal strFileOpenName As String, Optional ByVal varHeaderRow As Variant)
'Created by Robin G Brown, 5/9/97
'Read from a CSV file into the array
'Sub Declares
Const conSub = "FileOpen"
Dim lngCSVFileno As Long
Dim strCurrentCSVLine As String
Dim lngNumfields As Long
Dim booFirstLineRead As Boolean
Dim lngStartCommaPos As Long
Dim lngEndCommaPos As Long
Dim lngCSVLen As Long
    'Error Trap
    On Error GoTo FileOpen_ErrorHandler
    'Check the parameters to see if there is a header line in the data
    If IsMissing(varHeaderRow) Then
        flgTableFlags.FieldNames = False
    Else
        flgTableFlags.FieldNames = CBool(varHeaderRow)
    End If
    'Copy the FileOpen name to the internal variable
    strFileName = strFileOpenName
    'Open the file and read the first line
    lngCSVFileno = FreeFile()
    Open strFileName For Input Lock Read Write As #lngCSVFileno
    Line Input #lngCSVFileno, strCurrentCSVLine
    booFirstLineRead = True
    'Resize the array to the correct width
    lngNumfields = GetNumberOfCSVElements(strCurrentCSVLine)
    'lngRecordCount = 0
    ReDim strDataLines(1 To lngNumfields, 0 To lngRecordCount)
    If flgTableFlags.FieldNames = True Then
        'Read the header line into the array
        lngRecordCount = 0
        ReDim strDataLines(1 To lngNumfields, 0 To lngRecordCount)
        lngFieldCounter = 0
        lngStartCommaPos = 0
        For lngFieldCounter = 1 To lngNumfields
            lngEndCommaPos = InStr(lngStartCommaPos + 1, strCurrentCSVLine, conCSVTableSeparator)
            If lngEndCommaPos = 0 Then
                'Last word
                lngCSVLen = Len(strCurrentCSVLine)
                strDataLines(lngFieldCounter, lngRecordCount) = Trim$(Right$(strCurrentCSVLine, lngCSVLen - lngStartCommaPos))
            Else
                'Middle word
                strDataLines(lngFieldCounter, lngRecordCount) = Trim$(Mid$(strCurrentCSVLine, lngStartCommaPos + 1, lngEndCommaPos - lngStartCommaPos - 1))
            End If
            lngStartCommaPos = lngEndCommaPos
        Next
    Else
        'Read the first data line into the array
        lngRecordCount = 1
        ReDim strDataLines(1 To lngNumfields, 0 To lngRecordCount)
        lngFieldCounter = 1
        lngStartCommaPos = 0
        For lngFieldCounter = 1 To lngNumfields
            lngEndCommaPos = InStr(lngStartCommaPos + 1, strCurrentCSVLine, conCSVTableSeparator)
            If lngEndCommaPos = 0 Then
                'Last word
                lngCSVLen = Len(strCurrentCSVLine)
                strDataLines(lngFieldCounter, lngRecordCount) = Trim$(Right$(strCurrentCSVLine, lngCSVLen - lngStartCommaPos))
            Else
                'Middle word
                strDataLines(lngFieldCounter, lngRecordCount) = Trim$(Mid$(strCurrentCSVLine, lngStartCommaPos + 1, lngEndCommaPos - lngStartCommaPos - 1))
            End If
            lngStartCommaPos = lngEndCommaPos
        Next
    End If
    'Now process the rest of the lines in the file - assuming there are any of course!
    Do While Not EOF(lngCSVFileno)
        'Read the next line
        Line Input #lngCSVFileno, strCurrentCSVLine
        'Resize the array
        lngRecordCount = lngRecordCount + 1
        ReDim Preserve strDataLines(1 To lngNumfields, 0 To lngRecordCount)
        'Read the line into the array
        lngFieldCounter = 1
        lngStartCommaPos = 0
        For lngFieldCounter = 1 To lngNumfields
            lngEndCommaPos = InStr(lngStartCommaPos + 1, strCurrentCSVLine, conCSVTableSeparator)
            If lngEndCommaPos = 0 Then
                'Last word
                lngCSVLen = Len(strCurrentCSVLine)
                strDataLines(lngFieldCounter, lngRecordCount) = Trim$(Right$(strCurrentCSVLine, lngCSVLen - lngStartCommaPos))
            Else
                'Middle word
                strDataLines(lngFieldCounter, lngRecordCount) = Trim$(Mid$(strCurrentCSVLine, lngStartCommaPos + 1, lngEndCommaPos - lngStartCommaPos - 1))
            End If
            lngStartCommaPos = lngEndCommaPos
        Next
    Loop
    'Lastly close the file and clear any references
    Close #lngCSVFileno
Exit Sub
'Error Handler
FileOpen_ErrorHandler:
    Close #lngCSVFileno
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conClass & ":" & conSub, "Unexpected error : " & Err.Description
    End Select
    Exit Sub
End Sub

Public Sub FileSave()
'Created by Robin G Brown, 5/9/97
'Write the array away to a file
'Sub Declares
Const conSub = "FileSave"
Dim strFileText As String
Dim strLineText As String
Dim lngLinecounter As Long
Dim lngCSVFileno As Long
Dim lngLBound As Long
    'Error Trap
    On Error GoTo FileSave_ErrorHandler
    strFileText = ""
    If flgTableFlags.FieldNames = True Then
        lngLBound = 0
    Else
        lngLBound = 1
    End If
    For lngLinecounter = lngLBound To UBound(strDataLines, 2)
        strLineText = ""
        For lngFieldCounter = 1 To UBound(strDataLines, 1)
            strLineText = strLineText & strDataLines(lngFieldCounter, lngLinecounter) & conCSVTableSeparator
        Next
        'Trim last comma
        strLineText = Left$(strLineText, Len(strLineText) - 1)
        'Add to file
        strFileText = strFileText & strLineText & vbCrLf
    Next
    'Save
    lngCSVFileno = FreeFile()
    Open strFileName For Output Lock Read Write As #lngCSVFileno
    Print #lngCSVFileno, strFileText
    Close #lngCSVFileno
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

Public Sub AddRecord(Optional ByVal varPosition As Variant)
'Created by Robin G Brown, 5/9/97
'Add a Record to the end of the table and set the current line 'pointer' to that line _
    so that fieldvalues can be set
'Sub Declares
Const conSub = "AddRecord"
    'Error Trap
    On Error GoTo AddRecord_ErrorHandler
    lngRecordCount = lngRecordCount + 1
    ReDim Preserve strDataLines(1 To UBound(strDataLines, 1), 0 To lngRecordCount)
    lngCurrentLine = lngRecordCount
Exit Sub
'Error Handler
AddRecord_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conClass & ":" & conSub, "Unexpected error : " & Err.Description
    End Select
    Exit Sub
End Sub

'Public Sub SetSortOrderField(ByVal strSort As String)
''Created by Robin G Brown, 5/9/97
''Set which field(s) to sort the table on and re-order appropriately
''Sub Declares
'Const conSub = "SetSortOrderField"
'    'Error Trap
'    On Error GoTo SetSortOrderField_ErrorHandler
'    'This function is not supported yet
'    Err.Raise vbObjectError, conClass & ":" & conSub, "Unsupported function error"
'    'CODE_HERE
'Exit Sub
''Error Handler
'SetSortOrderField_ErrorHandler:
'    Select Case Err.Number
'    'Case ERROR_CODE_HERE
'        'ERROR_HANDLING_CODE_HERE
'    Case Else
'        Err.Raise Err.Number, conClass & ":" & conSub, "Unexpected error : " & Err.Description
'    End Select
'    Exit Sub
'End Sub

Public Sub MoveFirst()
'Created by Robin G Brown, 5/9/97
'Set lngCurrentLine to point at the first record
'Sub Declares
Const conSub = "MoveFirst"
    'Error Trap
    On Error Resume Next
    lngCurrentLine = 1
End Sub

Public Sub MoveNext()
'Created by Robin G Brown, 5/9/97
'Set lngCurrentLine to point at the next record, checking that this is allowed, _
    set the EOF property if we just moved to the last record
'Sub Declares
Const conSub = "MoveNext"
    'Error Trap
    On Error Resume Next
    If lngCurrentLine = lngRecordCount Then
        Err.Raise Err.Number, conClass & ":" & conSub, "Unexpected error : " & Err.Description
    Else
        lngCurrentLine = lngCurrentLine + 1
    End If
End Sub

Public Sub MoveLast()
'Created by Robin G Brown, 5/9/97
'Set lngCurrentLine to point at the last record
'Sub Declares
Const conSub = "MoveLast"
    'Error Trap
    On Error Resume Next
    lngCurrentLine = lngRecordCount
End Sub

Public Sub MovePrevious()
'Created by Robin G Brown, 5/9/97
'Set lngCurrentLine to point at the previous record
'Sub Declares
Const conSub = "MovePrevious"
    'Error Trap
    On Error Resume Next
    If lngCurrentLine = 1 Then
        Err.Raise Err.Number, conClass & ":" & conSub, "Unexpected error : " & Err.Description
    Else
        lngCurrentLine = lngCurrentLine - 1
    End If
End Sub

Public Sub FindFirst(ByVal strSearch As String)
'Created by Robin G Brown, 5/9/97
'Set lngCurrentLine to point at the first record that matches the search parameter
'the search parameter can only be a simple expression in the following form:
'   FIELDNAME =|>|< VALUE
'For the moment FIELDNAME can only be a number - RGB/8/9/97
'Sub Declares
Const conSub = "FindFirst"
Dim lngSearchField
Dim lngOperatorPosition As Long
Dim strSearchParameter As String
Dim lngSearchRecord As Long
    'Error Trap
    On Error GoTo FindFirst_ErrorHandler
    'Select the type of the search
    If InStr(strSearch, ">") > 0 Then
        'Greater than
        lngOperatorPosition = InStr(strSearch, ">")
        'lngSearchField = FindField(Left$(strSearch, lngOperatorPosition - 1))
        lngSearchField = Left$(strSearch, lngOperatorPosition - 1)
        strSearchParameter = Right$(strSearch, Len(strSearch) - lngOperatorPosition)
        'FindRecord does the actual work of searching
        Call FindRecord(conFindFirst, lngSearchField, ">", strSearchParameter)
    ElseIf InStr(strSearch, "<") > 0 Then
        'Less than
        lngOperatorPosition = InStr(strSearch, "<")
        'lngSearchField = FindField(Left$(strSearch, lngOperatorPosition - 1))
        lngSearchField = Left$(strSearch, lngOperatorPosition - 1)
        strSearchParameter = Right$(strSearch, Len(strSearch) - lngOperatorPosition)
        'FindRecord does the actual work of searching
        Call FindRecord(conFindFirst, lngSearchField, "<", strSearchParameter)
    Else
        '= is the default
        lngOperatorPosition = InStr(strSearch, "=")
        'lngSearchField = FindField(Left$(strSearch, lngOperatorPosition - 1))
        lngSearchField = Left$(strSearch, lngOperatorPosition - 1)
        strSearchParameter = Right$(strSearch, Len(strSearch) - lngOperatorPosition)
        'FindRecord does the actual work of searching
        Call FindRecord(conFindFirst, lngSearchField, "=", strSearchParameter)
    End If
Exit Sub
'Error Handler
FindFirst_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conClass & ":" & conSub, "Unexpected error : " & Err.Description
    End Select
    Exit Sub
End Sub

Public Sub FindNext(ByVal strSearch As String)
'Created by Robin G Brown, 5/9/97
'Set lngCurrentLine to point at the first record that matches the search parameter
'See FindFirst for a description of the function
'Sub Declares
Const conSub = "FindNext"
Dim lngSearchField
Dim lngOperatorPosition As Long
Dim strSearchParameter As String
Dim lngSearchRecord As Long
    'Error Trap
    On Error GoTo FindNext_ErrorHandler
    'Select the type of the search
    If InStr(strSearch, ">") > 0 Then
        'Greater than
        lngOperatorPosition = InStr(strSearch, ">")
        'lngSearchField = FindField(Left$(strSearch, lngOperatorPosition - 1))
        lngSearchField = Left$(strSearch, lngOperatorPosition - 1)
        strSearchParameter = Right$(strSearch, Len(strSearch) - lngOperatorPosition)
        'FindRecord does the actual work of searching
        Call FindRecord(conFindNext, lngSearchField, ">", strSearchParameter)
    ElseIf InStr(strSearch, "<") > 0 Then
        'Less than
        lngOperatorPosition = InStr(strSearch, "<")
        'lngSearchField = FindField(Left$(strSearch, lngOperatorPosition - 1))
        lngSearchField = Left$(strSearch, lngOperatorPosition - 1)
        strSearchParameter = Right$(strSearch, Len(strSearch) - lngOperatorPosition)
        'FindRecord does the actual work of searching
        Call FindRecord(conFindNext, lngSearchField, "<", strSearchParameter)
    Else
        '= is the default
        lngOperatorPosition = InStr(strSearch, "=")
        'lngSearchField = FindField(Left$(strSearch, lngOperatorPosition - 1))
        lngSearchField = Left$(strSearch, lngOperatorPosition - 1)
        strSearchParameter = Right$(strSearch, Len(strSearch) - lngOperatorPosition)
        'FindRecord does the actual work of searching
        Call FindRecord(conFindNext, lngSearchField, "=", strSearchParameter)
    End If
Exit Sub
'Error Handler
FindNext_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conClass & ":" & conSub, "Unexpected error : " & Err.Description
    End Select
    Exit Sub
End Sub

Public Sub FindLast(ByVal strSearch As String)
'Created by Robin G Brown, 5/9/97
'Set lngCurrentLine to point at the first record that matches the search parameter
'See FindFirst for a description of the function
'Sub Declares
Const conSub = "FindLast"
Dim lngSearchField
Dim lngOperatorPosition As Long
Dim strSearchParameter As String
Dim lngSearchRecord As Long
    'Error Trap
    On Error GoTo FindLast_ErrorHandler
    'Select the type of the search
    If InStr(strSearch, ">") > 0 Then
        'Greater than
        lngOperatorPosition = InStr(strSearch, ">")
        'lngSearchField = FindField(Left$(strSearch, lngOperatorPosition - 1))
        lngSearchField = Left$(strSearch, lngOperatorPosition - 1)
        strSearchParameter = Right$(strSearch, Len(strSearch) - lngOperatorPosition)
        'FindRecord does the actual work of searching
        Call FindRecord(conFindLast, lngSearchField, ">", strSearchParameter)
    ElseIf InStr(strSearch, "<") > 0 Then
        'Less than
        lngOperatorPosition = InStr(strSearch, "<")
        'lngSearchField = FindField(Left$(strSearch, lngOperatorPosition - 1))
        lngSearchField = Left$(strSearch, lngOperatorPosition - 1)
        strSearchParameter = Right$(strSearch, Len(strSearch) - lngOperatorPosition)
        'FindRecord does the actual work of searching
        Call FindRecord(conFindLast, lngSearchField, "<", strSearchParameter)
    Else
        '= is the default
        lngOperatorPosition = InStr(strSearch, "=")
        'lngSearchField = FindField(Left$(strSearch, lngOperatorPosition - 1))
        lngSearchField = Left$(strSearch, lngOperatorPosition - 1)
        strSearchParameter = Right$(strSearch, Len(strSearch) - lngOperatorPosition)
        'FindRecord does the actual work of searching
        Call FindRecord(conFindLast, lngSearchField, "=", strSearchParameter)
    End If
Exit Sub
'Error Handler
FindLast_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conClass & ":" & conSub, "Unexpected error : " & Err.Description
    End Select
    Exit Sub
End Sub

Public Sub FindPrevious(ByVal strSearch As String)
'Created by Robin G Brown, 5/9/97
'Set lngCurrentLine to point at the first record that matches the search parameter
'See FindFirst for a description of the function
'Sub Declares
Const conSub = "FindPrevious"
Dim lngSearchField
Dim lngOperatorPosition As Long
Dim strSearchParameter As String
Dim lngSearchRecord As Long
    'Error Trap
    On Error GoTo FindPrevious_ErrorHandler
    'Select the type of the search
    If InStr(strSearch, ">") > 0 Then
        'Greater than
        lngOperatorPosition = InStr(strSearch, ">")
        'lngSearchField = FindField(Left$(strSearch, lngOperatorPosition - 1))
        lngSearchField = Left$(strSearch, lngOperatorPosition - 1)
        strSearchParameter = Right$(strSearch, Len(strSearch) - lngOperatorPosition)
        'FindRecord does the actual work of searching
        Call FindRecord(conFindPrevious, lngSearchField, ">", strSearchParameter)
    ElseIf InStr(strSearch, "<") > 0 Then
        'Less than
        lngOperatorPosition = InStr(strSearch, "<")
        'lngSearchField = FindField(Left$(strSearch, lngOperatorPosition - 1))
        lngSearchField = Left$(strSearch, lngOperatorPosition - 1)
        strSearchParameter = Right$(strSearch, Len(strSearch) - lngOperatorPosition)
        'FindRecord does the actual work of searching
        Call FindRecord(conFindPrevious, lngSearchField, "<", strSearchParameter)
    Else
        '= is the default
        lngOperatorPosition = InStr(strSearch, "=")
        'lngSearchField = FindField(Left$(strSearch, lngOperatorPosition - 1))
        lngSearchField = Left$(strSearch, lngOperatorPosition - 1)
        strSearchParameter = Right$(strSearch, Len(strSearch) - lngOperatorPosition)
        'FindRecord does the actual work of searching
        Call FindRecord(conFindPrevious, lngSearchField, "=", strSearchParameter)
    End If
Exit Sub
'Error Handler
FindPrevious_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conClass & ":" & conSub, "Unexpected error : " & Err.Description
    End Select
    Exit Sub
End Sub

Public Sub DisplayInGrid(ByRef ctlGrid As Object)
'Created by Robin G Brown, 5/9/97
'Display the contents of the array in a grid, including headers and recordnumbers
'Sub Declares
Const conSub = "DisplayInGrid"
Dim lngLinecounter As Long
    'Error Trap
    On Error GoTo DisplayInGrid_ErrorHandler
    'Type of grid?
    'Don't check in order to reduce the number of references required
    'An error will be thrown if the wrong type of control is used
    With ctlGrid
        'Set size
        .Rows = UBound(strDataLines, 2) + 1
        .Cols = UBound(strDataLines, 1) + 1
        'Fill in headers, if they exist
        .Row = 0
        For lngFieldCounter = 1 To UBound(strDataLines, 1)
            .Col = lngFieldCounter
            If flgTableFlags.FieldNames = True Then
                .Text = strDataLines(lngFieldCounter, 0)
            Else
                .Text = lngFieldCounter
            End If
        Next
        'Set fixed cols/rows
        .FixedRows = 1
        .FixedCols = 1
        For lngLinecounter = 1 To UBound(strDataLines, 2)
            .Row = lngLinecounter
            .Col = 0
            .Text = lngLinecounter
            For lngFieldCounter = 1 To UBound(strDataLines, 1)
                .Col = lngFieldCounter
                .Text = strDataLines(lngFieldCounter, lngLinecounter)
            Next
        Next
    End With
Exit Sub
'Error Handler
DisplayInGrid_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conClass & ":" & conSub, "Unexpected error : " & Err.Description
    End Select
    Exit Sub
End Sub

'Public Sub SortTable()
''Created by Robin G Brown, 5/9/97
''Sort/Re-Sort the table by the sort fields
''Sub Declares
'Const conSub = "SortTable"
'    'Error Trap
'    On Error GoTo SortTable_ErrorHandler
'    'This function is not supported yet
'    Err.Raise vbObjectError, conClass & ":" & conSub, "Unsupported function error"
'    'CODE_HERE
'Exit Sub
''Error Handler
'SortTable_ErrorHandler:
'    Select Case Err.Number
'    'Case ERROR_CODE_HERE
'        'ERROR_HANDLING_CODE_HERE
'    Case Else
'        Err.Raise Err.Number, conClass & ":" & conSub, "Unexpected error : " & Err.Description
'    End Select
'    Exit Sub
'End Sub

'-----------------------------------------------------------------------------
'   Private Functions and Methods
'-----------------------------------------------------------------------------

Private Function FindField(ByVal strFieldName As String) As Long
'Created by Robin G Brown, 5/9/97
'Return the index of the given field
'Function Declares
Const conFunction = "FindField"
Dim lngFieldCounter As Long
    'Error Trap
    On Error GoTo FindField_ErrorHandler
    If flgTableFlags.FieldNames = False Then
        FindField = -1
        Err.Raise vbObjectError, conClass & ":" & conFunction, "Field Names not available."
        Exit Function
    Else
        For lngFieldCounter = 1 To UBound(strDataLines, 1)
            If LCase$(strDataLines(lngFieldCounter, 0)) = LCase$(strFieldName) Then
                FindField = lngFieldCounter
                Exit Function
            End If
        Next
        'If we get here then the field was not found
        FindField = -1
        Err.Raise vbObjectError, conClass & ":" & conFunction, "Field not found."
    End If
Exit Function
'Error Handler
FindField_ErrorHandler:
    FindField = -1
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conClass & ":" & conFunction, "Unexpected error : " & Err.Description
    End Select
    Exit Function
End Function

Private Sub FindRecord(ByVal bytDirection As Byte, ByVal lngSearchField As Long, ByVal strOperator As String, ByVal strParameter As String)
'Created by Robin G Brown, 8/9/97
'Find a record, if found setlngcurrentLine to the record found, _
    flgTableFlags.RecordFound to True and set flgTableFlags.BOF/EOF as appropriate
'Sub Declares
Const conSub = "FindRecord"
Dim lngSearchRecord As Long
Dim booSearchSucces As Boolean
Dim lngStartSearch As Long
Dim lngEndSearch As Long
Dim lngStep As Long
    'Error Trap
    On Error GoTo FindRecord_ErrorHandler
    'Set deafults
    Select Case bytDirection
    Case conFindNext
        lngStep = 1
        lngStartSearch = lngCurrentLine
        lngEndSearch = lngRecordCount
    Case conFindLast
        lngStep = -1
        lngStartSearch = lngRecordCount
        lngEndSearch = 1
    Case conFindPrevious
        lngStep = -1
        lngStartSearch = lngRecordCount
        lngEndSearch = lngCurrentLine
    Case Else
        'conFindFirst is deafult
        lngStep = 1
        lngStartSearch = 1
        lngEndSearch = lngRecordCount
    End Select
    'Start the search
    flgTableFlags.RecordFound = False
    Select Case strOperator
    Case ">"
        For lngSearchRecord = lngStartSearch To lngEndSearch Step lngStep
            If strDataLines(lngSearchField, lngSearchRecord) > strParameter Then
                flgTableFlags.RecordFound = True
                Exit For
            End If
        Next
    Case "<"
        For lngSearchRecord = lngStartSearch To lngEndSearch Step lngStep
            If strDataLines(lngSearchField, lngSearchRecord) < strParameter Then
                flgTableFlags.RecordFound = True
                Exit For
            End If
        Next
    Case Else
        For lngSearchRecord = lngStartSearch To lngEndSearch Step lngStep
            If strDataLines(lngSearchField, lngSearchRecord) = strParameter Then
                flgTableFlags.RecordFound = True
                Exit For
            End If
        Next
    End Select
Exit Sub
'Error Handler
FindRecord_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conClass & ":" & conSub, "Unexpected error : " & Err.Description
    End Select
    Exit Sub
End Sub

'-----------------------------------------------------------------------------
'   Properties
'-----------------------------------------------------------------------------

Property Get TableName() As String
'Created by Robin G Brown, 10/9/97
'Property Declares
Const conProperty = "TableName"
'Error Trap
    On Error Resume Next
    TableName = StripFileExtension(GetFileNameFromPath(strFileName))
End Property

Property Get RecordCount() As Long
'Created by Robin G Brown, 5/9/97
'Property Declares
Const conProperty = "RecordCount"
'Error Trap
    On Error Resume Next
    RecordCount = lngRecordCount
End Property

Property Get BeginningOfFile() As Boolean
'Created by Robin G Brown, 5/9/97
'Property Declares
Const conProperty = "BeginningOfFile"
'Error Trap
    On Error Resume Next
    If lngRecordCount = 0 Or lngCurrentLine = 1 Then
        BeginningOfFile = True
    Else
        BeginningOfFile = False
    End If
End Property

Property Get EndOfFile() As Boolean
'Created by Robin G Brown, 5/9/97
'Property Declares
Const conProperty = "EndOfFile"
'Error Trap
    On Error Resume Next
    If lngRecordCount = 0 Or lngCurrentLine = lngRecordCount Then
        EndOfFile = True
    Else
        EndOfFile = False
    End If
End Property

Property Get RecordFound() As String
'Created by Robin G Brown, 8/9/97
'Property Declares
Const conProperty = "RecordFound"
'Error Trap
    On Error Resume Next
    RecordFound = flgTableFlags.RecordFound
End Property

Property Get Bookmark() As Long
'Created by Robin G Brown, 5/9/97
'Return the value of lngCurrentLine to be saved externally as a 'BookMark'
'Property Declares
Const conProperty = "Bookmark"
'Error Trap
    On Error Resume Next
    Bookmark = lngCurrentLine
End Property

Property Let Bookmark(ByVal lngBookmark As Long)
'Created by Robin G Brown, 5/9/97
'Set lngCurrentLine equal to an external 'BookMark' value
'Property Declares
Const conProperty = "Bookmark"
    'Error Trap
    On Error Resume Next
    If lngBookmark < 1 Or lngBookmark > lngRecordCount Then
        Err.Raise vbObjectError, conClass & ":" & conProperty, "Invalid bookmark."
    Else
        lngCurrentLine = lngBookmark
    End If
End Property

Property Get FieldValue(ByVal varFieldIndex As Variant) As String
'Created by Robin G Brown, 5/9/97
'Return the value of the given field at the current record
'Property Declares
Const conFieldValue = "FieldValue"
Dim lngActualIndex As Long
'Error Trap
    On Error Resume Next
    If VarType(varFieldIndex) = vbString Then
        lngActualIndex = FindField(varFieldIndex)
    ElseIf IsNumeric(varFieldIndex) Then
        lngActualIndex = CLng(varFieldIndex)
    Else
        Err.Raise vbObjectError, conClass & ":" & conFieldValue, "Invalid field index"
    End If
    'Return the value
    Err.Clear
    FieldValue = strDataLines(lngActualIndex, lngCurrentLine)
    If Err.Number <> 0 Then
        Err.Raise Err.Number, conClass & ":" & conFieldValue, "Unexpected error : " & Err.Description
    End If
End Property

Property Let FieldValue(ByVal varFieldIndex As Variant, ByVal strNewValue As String)
'Created by Robin G Brown, 5/9/97
'Set the value of the given field at the current record
'Property Declares
Const conFieldValue = "FieldValue"
Dim lngActualIndex As Long
    'Error Trap
    On Error Resume Next
    If VarType(varFieldIndex) = vbString Then
        lngActualIndex = FindField(varFieldIndex)
    ElseIf IsNumeric(varFieldIndex) Then
        lngActualIndex = CLng(varFieldIndex)
    Else
        Err.Raise vbObjectError, conClass & ":" & conFieldValue, "Invalid field index"
    End If
    'Set the value
    Err.Clear
    strDataLines(lngActualIndex, lngCurrentLine) = strNewValue
    If Err.Number <> 0 Then
        Err.Raise Err.Number, conClass & ":" & conFieldValue, "Unexpected error : " & Err.Description
    End If
End Property

Property Get FieldNames() As String
'Created by Robin G Brown, 10/9/97
'Return the field names as a CSV string
'Property Declares
Const conProperty = "FieldNames"
Dim strFields As String
'Error Trap
    On Error Resume Next
    If flgTableFlags.FieldNames = True Then
        strFields = ""
        For lngFieldCounter = 1 To UBound(strDataLines, 1)
            strFields = strFields & strDataLines(lngFieldCounter, 0) & ","
        Next
        FieldNames = Left$(strFields, Len(strFields) - 1)
    Else
        Err.Raise vbObjectError, conClass & ":" & conProperty, "Field Names not available."
    End If
End Property

Property Let FieldNames(ByVal strNewFieldNames As String)
'Created by Robin G Brown, 10/9/97
'Read a CSV string to get the field names
'Property Declares
Const conProperty = "FieldNames"
Dim lngStartCommaPos As Long
Dim lngEndCommaPos As Long
Dim lngCSVLen As Long
    'Error Trap
    On Error GoTo FieldNames_ErrorHandler
    lngStartCommaPos = 0
    For lngFieldCounter = 1 To UBound(strDataLines, 1)
        lngEndCommaPos = InStr(lngStartCommaPos + 1, strNewFieldNames, conCSVTableSeparator)
        If lngEndCommaPos = 0 Then
            'Last word
            lngCSVLen = Len(strNewFieldNames)
            strDataLines(lngFieldCounter, 0) = Trim$(Right$(strNewFieldNames, lngCSVLen - lngStartCommaPos))
            Exit For
        Else
            'Middle word
            strDataLines(lngFieldCounter, 0) = Trim$(Mid$(strNewFieldNames, lngStartCommaPos + 1, lngEndCommaPos - lngStartCommaPos - 1))
        End If
        lngStartCommaPos = lngEndCommaPos
    Next
    flgTableFlags.FieldNames = True
Exit Property
'Error Handler
FieldNames_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conClass & ":" & conProperty, "Unexpected error : " & Err.Description
    End Select
    Exit Property
End Property

Private Sub Class_Terminate()
'Created by Robin G Brown, 5/9/97
'Sub Declares
    'Error Trap
    On Error Resume Next
    'Clear the arrays
    ReDim strDataLines(0 To 0, 0 To 0)
    ReDim strDataLines(0 To 0)
End Sub


