Attribute VB_Name = "modCSVLib"
'Created by Robin G Brown, 5th September 1997
'-----------------------------------------------------------------------------
'   This module contains code for _
    handling CSV data
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "modCSVLib"
Private Const conCSVSeparator = ","
Private lngCounter As Long
Private lngReturn As Long

Public Function GetNumberOfCSVElements(ByVal strCSVSearch As String) As Long
'Created by Robin G Brown, 5/9/97
'Find the number of elements in a CSV string
'Function Declares
Const conFunction = "GetNumberOfCSVElements"
Dim strChar As String * 1
Dim lngSearchPos As String
Dim lngCharsFound As Long
    'Error Trap
    On Error GoTo GetNumberOfCSVElements_ErrorHandler
    lngCharsFound = 0
    lngSearchPos = InStr(1, strCSVSearch, conCSVSeparator)
    Do While lngSearchPos > 0
        lngCharsFound = lngCharsFound + 1
        lngSearchPos = InStr(lngSearchPos + 1, strCSVSearch, conCSVSeparator)
    Loop
    GetNumberOfCSVElements = lngCharsFound + 1
Exit Function
'Error Handler
GetNumberOfCSVElements_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conModule & ":" & conFunction, "Unexpected error : " & Err.Description
    End Select
    GetNumberOfCSVElements = -1
    Exit Function
End Function

Public Function WriteArrayToCSV(ByRef strArray() As String) As String
'Created by Robin G Brown, 20/5/97
'Created a CSV string from an array
'Function Declares
Const conFunction = "WriteArrayToCSV"
Dim lngArraycounter As Long
Dim strCSV As String
    'Error Trap
    On Error GoTo WriteArrayToCSV_ErrorHandler
    strCSV = ""
    For lngArraycounter = LBound(strArray, 1) To UBound(strArray, 1)
        strCSV = strCSV & strArray(lngArraycounter) & ","
    Next
    strCSV = Left$(strCSV, Len(strCSV) - 1)
    WriteArrayToCSV = strCSV
Exit Function
'Error Handler
WriteArrayToCSV_ErrorHandler:
    WriteArrayToCSV = ""
    Exit Function
End Function

Public Function WriteCSVToArray(ByVal strCSV As String, ByRef strArray() As String) As Boolean
'Created by Robin G Brown
'Writes the individual elements of strCSV string to a 1 dimensional array
Dim lngXCommaPos As Long
Dim lngYCommaPos As Long
Dim lngZLen As Long
Dim lngLBound As Long
Dim lngUBound As Long
    On Error GoTo WriteCSVToArray_ErrorHandler
    lngLBound = LBound(strArray, 1)
    lngUBound = UBound(strArray, 1)
    lngXCommaPos = InStr(strCSV, ",")
    lngZLen = Len(strCSV)
    If lngXCommaPos = 0 Then
        'Assume there's only one word
        strArray(lngLBound) = Trim$(strCSV)
        WriteCSVToArray = True
        Exit Function
    End If
    lngCounter = lngLBound
    lngXCommaPos = 0
    Do
        lngYCommaPos = InStr(lngXCommaPos + 1, strCSV, ",")
        If lngYCommaPos = 0 Then
            'Last word
            strArray(lngCounter) = Trim$(Right$(strCSV, lngZLen - lngXCommaPos))
            Exit Do
        Else
            'Middle word
            strArray(lngCounter) = Trim$(Mid$(strCSV, lngXCommaPos + 1, lngYCommaPos - lngXCommaPos - 1))
            lngCounter = lngCounter + 1
        End If
        lngXCommaPos = lngYCommaPos
    Loop While lngXCommaPos > 0 And lngCounter <= lngUBound
    WriteCSVToArray = True
    Exit Function
'Error Handler
WriteCSVToArray_ErrorHandler:
    WriteCSVToArray = False
    Exit Function
End Function

Public Function WriteCSVToUnsizedArray(ByVal strCSV As String, ByRef strArray() As String) As Boolean
'Created by Robin G Brown, 6/8/97
'Writes the individual elements of strCSV string to a 1 dimensional array, resizing as we go
Dim lngXCommaPos As Long
Dim lngYCommaPos As Long
Dim lngZLen As Long
    On Error GoTo WriteCSVToUnsizedArray_ErrorHandler
    lngXCommaPos = InStr(strCSV, ",")
    lngZLen = Len(strCSV)
    'Resize the array to get rid of any previous data
    ReDim strArray(1 To 1)
    If lngXCommaPos = 0 Then
        'Assume there's only one word
        strArray(1) = Trim$(strCSV)
        WriteCSVToUnsizedArray = True
        Exit Function
    End If
    lngCounter = 1
    lngXCommaPos = 0
    Do
        lngYCommaPos = InStr(lngXCommaPos + 1, strCSV, ",")
        ReDim Preserve strArray(1 To lngCounter)
        If lngYCommaPos = 0 Then
            'Last word
            strArray(lngCounter) = Trim$(Right$(strCSV, lngZLen - lngXCommaPos))
            Exit Do
        Else
            'Middle word
            strArray(lngCounter) = Trim$(Mid$(strCSV, lngXCommaPos + 1, lngYCommaPos - lngXCommaPos - 1))
            lngCounter = lngCounter + 1
        End If
        lngXCommaPos = lngYCommaPos
    Loop While lngXCommaPos > 0
    WriteCSVToUnsizedArray = True
    Exit Function
'Error Handler
WriteCSVToUnsizedArray_ErrorHandler:
    WriteCSVToUnsizedArray = False
    Exit Function
End Function


