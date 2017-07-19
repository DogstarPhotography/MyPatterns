VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Created by Robin G Brown, 30th July 1997
'-----------------------------------------------------------------------------
'   This class module contains code for _
    handling files
'   Use srv as a prefix for this class
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "Server"
'Private properties
Private lngCounter As Long
Private lngReturn As Long
Private strSourceDir As String
Private strDestDir As String

Public Sub CopyFile(ByVal strSourceFile As String, ByVal strDestinationFile As String)
'Created by Robin G Brown, 30/7/97
'CopyFile a file from source to destination
'Sub Declares
Const conSub = "CopyFile"
Dim strPath As String
    'Error Trap
    On Error GoTo CopyFile_ErrorHandler
    strPath = GetPathFromPathAndFilename(strSourceFile)
    If strPath = "" Then
        strSourceFile = strSourceDir & "\" & strSourceFile
    End If
    strPath = GetPathFromPathAndFilename(strDestinationFile)
    If strPath = "" Then
        strDestinationFile = strDestDir & "\" & strDestinationFile
    End If
    FileCopy strSourceFile, strDestinationFile
Exit Sub
'Error Handler
CopyFile_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise vbObjectError, conModule & ":" & conSub, "Unexpected error : " & Err.Description
    End Select
    Exit Sub
End Sub

Public Sub DeleteFile(ByVal strFile As String)
'Created by Robin G Brown, 30/7/97
'Delete the given file
'Sub Declares
Const conSub = "DeleteFile"
Dim strPath As String
    'Error Trap
    On Error GoTo DeleteFile_ErrorHandler
    strPath = GetPathFromPathAndFilename(strFile)
    If strPath = "" Then
        strFile = strSourceDir & "\" & strFile
    End If
    Kill strFile
Exit Sub
'Error Handler
DeleteFile_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise vbObjectError, conModule & ":" & conSub, "Unexpected error : " & Err.Description
    End Select
    Exit Sub
End Sub

Public Sub MoveFile(ByVal strSourceFile As String, ByVal strDestinationFile As String)
'Created by Robin G Brown, 30/7/97
'Move a file from source to destination
'Sub Declares
Const conSub = "MoveFile"
    'Error Trap
    On Error GoTo MoveFile_ErrorHandler
    Call CopyFile(strSourceFile, strDestinationFile)
    Call DeleteFile(strSourceFile)
Exit Sub
'Error Handler
MoveFile_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise vbObjectError, conModule & ":" & conSub, "Unexpected error : " & Err.Description
    End Select
    Exit Sub
End Sub

Public Sub RenameFile(ByVal strFile As String, ByVal strRename As String)
'Created by Robin G Brown, 30/7/97
'Rename a file
'Sub Declares
Const conSub = "RenameFile"
Dim strPath As String
    'Error Trap
    On Error GoTo RenameFile_ErrorHandler
    strPath = GetPathFromPathAndFilename(strFile)
    If strPath = "" Then
        strFile = strSourceDir & "\" & strFile
    End If
    Name strFile As strRename
Exit Sub
'Error Handler
RenameFile_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise vbObjectError, conModule & ":" & conSub, "Unexpected error : " & Err.Description
    End Select
    Exit Sub
End Sub

Public Sub CreateDirectory(ByVal strNewDirectory As String)
'Created by Robin G Brown, 1/8/97
'Create a directory, in the source directory if no path is specified
'Sub Declares
Const conSub = "CreateDirectory"
    'Error Trap
    On Error GoTo CreateDirectory_ErrorHandler
    If InStr(strNewDirectory, "\") > 0 Then
        MkDir strNewDirectory
    Else
        MkDir strSourceDir & "\" & strNewDirectory
    End If
Exit Sub
'Error Handler
CreateDirectory_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise vbObjectError, conModule & ":" & conSub, "Unexpected error : " & Err.Description
    End Select
    Exit Sub
End Sub

Public Sub DeleteDirectory(ByVal strDirectory As String)
'Created by Robin G Brown, 1/8/97
'Create a directory, in the source directory if no path is specified
'Sub Declares
Const conSub = "DeleteDirectory"
    'Error Trap
    On Error GoTo DeleteDirectory_ErrorHandler
    If InStr(strDirectory, "\") > 0 Then
        'Remove all files in strdirectory
        Kill strDirectory & "\*.*"
        RmDir strDirectory
    Else
        Kill strSourceDir & "\" & strDirectory & "\*.*"
        RmDir strSourceDir & "\" & strDirectory
    End If
Exit Sub
'Error Handler
DeleteDirectory_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise vbObjectError, conModule & ":" & conSub, "Unexpected error : " & Err.Description
    End Select
    Exit Sub
End Sub

Public Sub RenameDirectory(ByVal strDirectory As String, ByVal strRename As String)
'Created by Robin G Brown, 1/8/97
'Rename a directory
'Sub Declares
Const conSub = "RenameDirectory"
Dim strPath As String
    'Error Trap
    On Error GoTo RenameDirectory_ErrorHandler
    strPath = GetPathFromPathAndFilename(strDirectory)
    If strPath = "" Then
        strDirectory = strSourceDir & "\" & strDirectory
        strRename = strSourceDir & "\" & strDirectory
    Else
        strRename = strPath & "\" & strDirectory
    End If
    Name strDirectory As strRename
Exit Sub
'Error Handler
RenameDirectory_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise vbObjectError, conModule & ":" & conSub, "Unexpected error : " & Err.Description
    End Select
    Exit Sub
End Sub

Property Get Source() As String
'Created by Robin G Brown, 30/7/97
'Sub Declares
    'Error Trap
    On Error Resume Next
    Source = strSourceDir
End Property

Property Let Source(ByVal strDirectory As String)
'Created by Robin G Brown, 30/7/97
'Sub Declares
    'Error Trap
    On Error Resume Next
    If Right$(strDirectory, 1) = "\" Then
        strDirectory = Left$(strDirectory, Len(strDirectory) - 1)
    End If
    strSourceDir = strDirectory
End Property

Property Get Destination() As String
'Created by Robin G Brown, 30/7/97
'Sub Declares
    'Error Trap
    On Error Resume Next
    Destination = strDestDir
End Property

Property Let Destination(ByVal strDirectory As String)
'Created by Robin G Brown, 30/7/97
'Sub Declares
    'Error Trap
    On Error Resume Next
    If Right$(strDirectory, 1) = "\" Then
        strDirectory = Left$(strDirectory, Len(strDirectory) - 1)
    End If
    strDestDir = strDirectory
End Property


