Attribute VB_Name = "modFileLib"
'Created by Robin G Brown, 19th May 1997
'-----------------------------------------------------------------------------
'   This module contains code for _
    handling file access and creating and using a log file
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "modFileLib"
Private intCounter As Integer
Private intReturn As Integer
'Log file declares
Private booLoggingFlag As Boolean
Private lngLogFileNumber As Long
'Filter Constants
Public Const conAnyFilter = "All Files (*.*)|*.*"
Public Const conTXTFilter = "Text Files (*.txt)|*.txt"
Public Const conCSVFilter = "CSV Files (*.csv)|*.csv"
Public Const conLogFilter = "Log Files (*.log)|*.log"
Public Const conWordFilter = "Word Files (*.doc)|*.doc"
Public Const conExcelFilter = "Excel Files (*.xl*)|*.xl*"
Public Const conAccessFilter = "Access Databases (*.mdb)|*.mdb"
Public Const conHTMLFilter = "HTML Files (*.htm)|*.htm"
Public Const conBATFilter = "Batch Files (*.bat)|*.bat"
Public Const conEXEFilter = "Programs (*.exe)|*.exe"
Public Const conAnyTXTFilter = conTXTFilter & "|" & conAnyFilter
Public Const conAnyCSVFilter = conCSVFilter & "|" & conAnyFilter
Public Const conOfficeFilter = conWordFilter & "|" & conExcelFilter & "|" & conAnyFilter
Public Const conAnyAccessFilter = conAccessFilter & "|" & conAnyFilter
Public Const conDataFilter = conCSVFilter & "|" & conTXTFilter & "|" & conAnyFilter
Public Const conAnyHTMLFilter = conHTMLFilter & "|" & conAnyFilter

'Public Function CommonDialogFunction(ByRef cdlFileDialog As CommonDialog) As String
''Created by Robin G Brown, INSERT_DATE
''DESCRIPTION_OF_SUBROUTINE_FUNCTIONALITY_HERE
''Sub Declares
'Const conSub = "CommonDialogFunction"
'    'Error Trap
'    On Error GoTo CommonDialogFunction_ErrorHandler
'    With cdlFileDialog
'        .CancelError = True
'        .Filter = conAnyFilter
'        '.ShowOpen
'        '.ShowSave
'        CommonDialogFunction = .filename
'    End With
'Exit Function
''Error Handler
'CommonDialogFunction_ErrorHandler:
'    Select Case Err.Number
'    Case cdlCancel
'        'Cancel pressed, ignore
'    Case Else
'        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
'    End Select
'    CommonDialogFunction = ""
'    Exit Function
'End Function

Public Function GetKeyValue(ByVal strFileName As String, ByVal strSection As String, ByVal strKey As String) As String
'Created by Robin G Brown, 24/7/97
'Retrieve a Key from an INI style file
'e.g.
'   [SECTION]
'   Key=RETURNED_VALUE_HERE
'Function Declares
Const conFunction = "GetKeyValue"
Dim strLine As String
Dim lngFileNo As Long
Dim intAssignPos As Integer
    'Error Trap
    On Error GoTo GetKeyValue_ErrorHandler
    lngFileNo = FreeFile()
    Open strFileName For Input As #lngFileNo
    Do While Not EOF(lngFileNo)
        Line Input #lngFileNo, strLine
        If InStr(strLine, "[") > 0 _
        And InStr(strLine, strSection) > 0 _
        And InStr(strLine, "]") > 0 Then
            'Search the section for the Key
            Do While Not EOF(lngFileNo)
                Line Input #lngFileNo, strLine
                If InStr(strLine, strKey) Then
                    intAssignPos = InStr(strLine, "=")
                    GetKeyValue = Right$(strLine, Len(strLine) - intAssignPos)
                    Exit Function
                End If
            Loop
            Exit Do
        End If
        Line Input #lngFileNo, strLine
    Loop
    Close #lngFileNo
    GetKeyValue = ""
Exit Function
'Error Handler
GetKeyValue_ErrorHandler:
    Select Case Err.Number
    Case 62
        'Line input past EOF - ignore & exit
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conFunction
    End Select
    GetKeyValue = ""
    Exit Function
End Function

Public Function GetFileNameFromPath(ByVal strPath As String) As String
'Created by Robin G Brown
'Extracts the file name from a string, returns "" On Error
'Function Declares
Const conFunction = "GetFileNameFromPath"
Dim intCharPos As Integer
Dim strSearch As String
    'Error Trap
    On Error GoTo GetFileNameFromPath_ErrorHandler
    GetFileNameFromPath = ""
    strSearch = strPath
    intCharPos = FileGetLastCharPos(strSearch, "\")
    If intCharPos > 0 Then
        GetFileNameFromPath = Right$(strSearch, Len(strSearch) - intCharPos)
    Else
        GetFileNameFromPath = ""
    End If
Exit Function
'Error Handler
GetFileNameFromPath_ErrorHandler:
    GetFileNameFromPath = ""
    Exit Function
End Function

Public Function GetFileExtension(ByVal strFileName As String) As String
'Created by Robin G Brown
'Extracts the file extension from a string, returns "" On Error
'Function Declares
Const conFunction = "GetFileExtension"
Dim intCharPos As Integer
Dim strSearch As String
    'Error Trap
    On Error GoTo GetFileExtension_ErrorHandler
    GetFileExtension = ""
    strSearch = strFileName
    intCharPos = FileGetLastCharPos(strSearch, ".")
    If intCharPos > 0 Then
        GetFileExtension = Right$(strSearch, Len(strSearch) - intCharPos)
    Else
        GetFileExtension = ""
    End If
Exit Function
'Error Handler
GetFileExtension_ErrorHandler:
    GetFileExtension = ""
    Exit Function
End Function

Public Function StripFileExtension(ByVal strFileName As String) As String
'Created by Robin G Brown
'Extracts the file name without the extension from a string, returns "" On Error
'Function Declares
Const conFunction = "StripFileExtension"
Dim intCharPos As Integer
Dim strSearch As String
    'Error Trap
    On Error GoTo StripFileExtension_ErrorHandler
    StripFileExtension = ""
    strSearch = strFileName
    intCharPos = FileGetLastCharPos(strSearch, ".")
    If intCharPos > 0 Then
        StripFileExtension = Left$(strSearch, intCharPos - 1)
    Else
        StripFileExtension = ""
    End If
Exit Function
'Error Handler
StripFileExtension_ErrorHandler:
    StripFileExtension = ""
    Exit Function
End Function

Public Function IsValidPath(ByVal strPathName As String) As Boolean
'Created by Robin G Brown
'Checks as to wether a string is a valid path
'Function Declares
Const conFunction = "IsValidPath"
    'Error Trap
    On Error GoTo IsValidPath_ErrorHandler
    If InStr(strPathName, "/") > 0 _
    Or InStr(strPathName, "*") > 0 _
    Or InStr(strPathName, "?") > 0 _
    Or InStr(strPathName, Chr$(34)) > 0 _
    Or InStr(strPathName, "<") > 0 _
    Or InStr(strPathName, ">") > 0 _
    Or InStr(strPathName, "|") > 0 _
    Then
        IsValidPath = False
        Exit Function
    End If
    IsValidPath = True
On Error GoTo 0
Exit Function
'Error Handler
IsValidPath_ErrorHandler:
    IsValidPath = False
    Exit Function
End Function

Public Function IsValidLongFileName(ByVal strFileName As String) As Boolean
'Created by Robin G Brown
'Checks as to wether a string is a valid long filename
'Function Declares
Const conFunction = "IsValidLongFileName"
    'Error Trap
    On Error GoTo IsValidLongFileName_ErrorHandler
    If Len(strFileName) > 255 Then
        IsValidLongFileName = False
        Exit Function
    End If
    If InStr(strFileName, "\") > 0 _
    Or InStr(strFileName, "/") > 0 _
    Or InStr(strFileName, ":") > 0 _
    Or InStr(strFileName, "*") > 0 _
    Or InStr(strFileName, "?") > 0 _
    Or InStr(strFileName, Chr$(34)) > 0 _
    Or InStr(strFileName, "<") > 0 _
    Or InStr(strFileName, ">") > 0 _
    Or InStr(strFileName, "|") > 0 _
    Then
        IsValidLongFileName = False
        Exit Function
    End If
    IsValidLongFileName = True
On Error GoTo 0
Exit Function
'Error Handler
IsValidLongFileName_ErrorHandler:
    IsValidLongFileName = False
    Exit Function
End Function

Public Function Is83FileName(ByVal strFileName As String) As Boolean
'Created by Robin G Brown
'Checks as to wether a string is a valid 8.3 filename
'Function Declares
Const conFunction = "Is83FileName"
    'Error Trap
    On Error GoTo Is83FileName_ErrorHandler
    If Len(strFileName) > 12 Then
        Is83FileName = False
        Exit Function
    End If
    If InStr(strFileName, ".") = 0 Then
        'This is not 100% accurate! - RGB
        Is83FileName = False
        Exit Function
    End If
    If InStr(strFileName, "\") > 0 _
    Or InStr(strFileName, "/") > 0 _
    Or InStr(strFileName, ":") > 0 _
    Or InStr(strFileName, "*") > 0 _
    Or InStr(strFileName, "?") > 0 _
    Or InStr(strFileName, Chr$(34)) > 0 _
    Or InStr(strFileName, "<") > 0 _
    Or InStr(strFileName, ">") > 0 _
    Or InStr(strFileName, "|") > 0 _
    Or InStr(strFileName, " ") > 0 _
    Then
        Is83FileName = False
        Exit Function
    End If
    Is83FileName = True
On Error GoTo 0
Exit Function
'Error Handler
Is83FileName_ErrorHandler:
    Is83FileName = False
    Exit Function
End Function

Public Function CheckForDuplicate(ByVal strFileAndPath As String) As Boolean
'Created by Robin G Brown
'Checks for a duplicate file, _
    asks user for action, _
    returns true of overwrite is OK, false otherwise
'Function Declares
Const conFunction = "CheckForDuplicate"
Dim strMatch As String
Dim strFile As String
Dim intReply As Integer
    'Error Trap
    On Error GoTo CheckForDuplicate_ErrorHandler
    strFile = GetFileNameFromPath(strFileAndPath)
    strMatch = Dir(strFileAndPath)
    If strMatch = strFile Then
        intReply = MsgBox("Duplicate file found '" & strFile & "', overwrite?", vbExclamation + vbOKCancel, App.ProductName)
        If intReply = vbCancel Then
            CheckForDuplicate = False
        Else
            CheckForDuplicate = True
        End If
    Else
        CheckForDuplicate = True
    End If
On Error GoTo 0
Exit Function
'Error Handler
CheckForDuplicate_ErrorHandler:
    MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conFunction
    CheckForDuplicate = False
    Exit Function
End Function

Public Function DeleteFile(ByVal strFileName As String) As Integer
'By Mark M, 3/11/1995
'THis will try to delete the specified file, and return False if there is an error
    On Error GoTo DeleteFile_ErrorHandler
    If CheckForDuplicate(strFileName) = True Then
        Kill strFileName
        DeleteFile = True
    Else
        DeleteFile = False
    End If
Exit Function
DeleteFile_ErrorHandler:
    DeleteFile = False
    Exit Function
End Function

Public Function CopyFiles(strThisFileName As String, strNewFileName As String) As Integer
'by Mark M, 2/11/95
'to copy one file to another.
    On Error GoTo CopyFiles_Errorhandler
    CopyFiles = True
    FileCopy strThisFileName, strNewFileName
    Exit Function
CopyFiles_Errorhandler:
    CopyFiles = False
    Exit Function
End Function

Public Function GetPathFromPathAndFilename(ByVal strPathAndFilename As String) As String
'Created by Robin G Brown
'Extracts the path (excluding last '\') from a string, _
    returns "" On Error or if there is no path
'Function Declares
Const conFunction = "GetPathFromPathAndFilename"
Dim intCharPos As Integer
Dim strSearch As String
    'Error Trap
    On Error GoTo GetPathFromPathAndFilename_ErrorHandler
    GetPathFromPathAndFilename = ""
    strSearch = strPathAndFilename
    intCharPos = FileGetLastCharPos(strSearch, "\")
    If intCharPos > 0 Then
        strSearch = Left$(strSearch, intCharPos - 1)
    End If
    GetPathFromPathAndFilename = strSearch
Exit Function
'Error Handler
GetPathFromPathAndFilename_ErrorHandler:
    On Error Resume Next
    GetPathFromPathAndFilename = ""
    Exit Function
End Function

'This is a copy of GetLastCharPos from the String Handling Library
'    renamed for compatibility purposes
Private Function FileGetLastCharPos(ByVal strSourceField As String, ByVal strCharacter As String) As Integer
'Created by Robin G Brown
'Takes a string and finds the position of the last given character in it, _
    will return 0 if character is not found or if more than one character is entered _
    or if there is an error
'Function Declares
Const conFunction = "FileGetLastCharPos"
Dim intAnother As Integer
Dim intStart As Integer
Dim intLast As Integer
    'Error Trap
    On Error GoTo FileGetLastCharPos_ErrorHandler
    FileGetLastCharPos = 0
    If Len(strCharacter) > 1 Then
        Exit Function
    End If
    'find first position
    intStart = 1
    intAnother = InStr(intStart, strSourceField, strCharacter, 1)
    intLast = intAnother
    Do While intAnother >= intStart
        intLast = intAnother
        intStart = intAnother + 1
        intAnother = InStr(intStart, strSourceField, strCharacter, 1)
    Loop
    FileGetLastCharPos = intLast
Exit Function
'Error Handler
FileGetLastCharPos_ErrorHandler:
    FileGetLastCharPos = 0
    Exit Function
End Function

Public Sub OpenLogfile(ByVal strFileName As String)
'Created by Robin G Brown, 21/7/97
'Get the name of the log file and open it for output
'Sub Declares
    'Error Trap
    On Error Resume Next
    lngLogFileNumber = FreeFile()
    Open strFileName For Output As lngLogFileNumber
    booLoggingFlag = True
    WriteLog "Log file opened."
End Sub

Public Sub AppendLogfile(ByVal strFileName As String)
'Created by Robin G Brown, 21/7/97
'Get the name of the log file and open it for append
'Sub Declares
    'Error Trap
    On Error Resume Next
    lngLogFileNumber = FreeFile()
    Open strFileName For Append As lngLogFileNumber
    booLoggingFlag = True
    WriteLog "Log file opened for append."
End Sub

Public Sub WriteLog(ByVal strLogItem As String)
'Created by Robin G Brown, 21/7/97
'Write a line to the logfile
'Sub Declares
    'Error Trap
    On Error Resume Next
    'Check to see if logging is required
    If booLoggingFlag = True Then
        'Write a line to the file with the time
        Write #lngLogFileNumber, Now, strLogItem
    End If
End Sub

Public Sub CloseLogfile()
'Created by Robin G Brown, 21/7/97
'Switch the logging off
'Sub Declares
    'Error Trap
    On Error Resume Next
    WriteLog "Log file closed."
    Close lngLogFileNumber
    lngLogFileNumber = 0
    booLoggingFlag = False
End Sub






