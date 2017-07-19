VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SecFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Created by Robin G Brown, 14th august 1997
'-----------------------------------------------------------------------------
'   This class module contains code for _
    implementing a sections object using aggregation
'   Use sff as a prefix for this class
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conClass = "SecFile"
Private secFileSections() As Section
Private strFilename As String

Private Sub Class_Initialize()
'Created by Robin G Brown, 14/8/97
'Sub Declares
    'Error Trap
    On Error Resume Next
    ReDim secFileSections(0 To 0)
    If Err.Number <> 0 Then
        Err.Raise Err.Number, conClass & ":Initialize", "Unexpected error : " & Err.Description
    End If
End Sub

Property Get Sections() As String
'Created by Robin G Brown, 18/8/97
'Return the number of sections there are in this object
'Property Declares
Const conProperty = "Sections"
'Error Trap
    On Error Resume Next
    Sections = UBound(secFileSections, 1)
End Property

Property Get IndexSection(ByVal lngIndex As Long) As String
'Created by Robin G Brown, 18/8/97
'Return the Action name for a given index
'Sub Declares
Const conProperty = "IndexSection"
    'Error Trap
    On Error Resume Next
    IndexSection = secFileSections(lngIndex).Name
End Property

Public Sub FileOpen(ByVal strNewFileName As String)
'Created by Robin G Brown, 14/8/97
'Read from a file and create Sections and keys based on the file information
'Sub Declares
Const conSub = "FileOpen"
Dim strFileLine As String
Dim lngFileNumber As Long
Dim lngAssignPosition As Long
Dim strSectionPart As String
Dim secCurrent As Section
Dim strKeyPart As String
Dim strValuePart As String
Dim lngUbound As Long
    'Error Trap
    On Error GoTo FileOpen_ErrorHandler
    strFilename = strNewFileName
    lngFileNumber = FreeFile()
    Open strFilename For Input As #lngFileNumber
    'Redim the sections array to prevent multiple calls to FileOpen _
        from appending to the end of the array! - RGB/11/9/97
    ReDim secFileSections(0 To 0)
    'Now process the file
    Do While Not EOF(lngFileNumber)
        Line Input #lngFileNumber, strFileLine
        'Read line for data
        If InStr(strFileLine, "[") > 0 Then
            lngUbound = UBound(secFileSections, 1)
            'If secCurrent is not nothing then set the previous section to be secCurrent
            If Not secCurrent Is Nothing Then
                Set secFileSections(lngUbound) = secCurrent
            End If
            'Create a new section in the array
            If lngUbound = 0 Then
                ReDim secFileSections(1 To lngUbound + 1)
            Else
                ReDim Preserve secFileSections(1 To lngUbound + 1)
            End If
            Set secCurrent = secFileSections(lngUbound + 1)
            'Then create a new section
            strSectionPart = Mid$(strFileLine, 2, Len(strFileLine) - 2)
            Set secCurrent = New Section
            secCurrent.Name = strSectionPart
        ElseIf InStr(strFileLine, "=") > 0 Then
            If Not secCurrent Is Nothing Then
                'Only add keys if there is a current section
                lngAssignPosition = InStr(strFileLine, "=")
                strKeyPart = Left$(strFileLine, lngAssignPosition - 1)
                strValuePart = Right$(strFileLine, Len(strFileLine) - lngAssignPosition)
                'Add to the current section
                secCurrent.KeyValue(strKeyPart) = strValuePart
            End If
        Else
            'Not a valid line, so ignore
        End If
    Loop
    'Add the last section, if there is one
    lngUbound = UBound(secFileSections, 1)
    If Not secCurrent Is Nothing Then
        Set secFileSections(lngUbound) = secCurrent
    End If
    Close #lngFileNumber
Exit Sub
'Error Handler
FileOpen_ErrorHandler:
    Close #lngFileNumber
    Err.Raise Err.Number, conClass & ":" & conSub, "Unexpected error : " & Err.Description
    Exit Sub
End Sub

Public Sub FileSave(ByVal strNewFileName As String)
'Created by Robin G Brown, 28/8/97
'Output all Sections and keys to a file
'Sub Declares
Const conSub = "FileSave"
Dim strConcatenated As String
Dim lngFileNumber As Long
Dim lngSectionCounter As Long
Dim lngKeyCounter As Long
    'Error Trap
    On Error Resume Next
    If strNewFileName = "" Then
        Err.Raise vbObjectError, conClass & ":" & conSub, "Invalid filename."
    End If
    strFilename = strNewFileName
    strConcatenated = ""
    For lngSectionCounter = 1 To UBound(secFileSections, 1)
        With secFileSections(lngSectionCounter)
            'Add [SECTION]
            strConcatenated = strConcatenated & "[" & .Name & "]" & vbCrLf
            For lngKeyCounter = 1 To .KeyCount
                'Add KEY=VALUE
                strConcatenated = strConcatenated & .IndexKey(lngKeyCounter) & "=" & .KeyValue(.IndexKey(lngKeyCounter)) & vbCrLf
            Next
        End With
        'Add a blank line
        strConcatenated = strConcatenated & vbCrLf
    Next
    lngFileNumber = FreeFile()
    Open strFilename For Output As #lngFileNumber
    Print #lngFileNumber, strConcatenated
    Close #lngFileNumber
End Sub

'Property Get File() As String
''Created by Robin G Brown, 14/8/97
''Output all Sections and keys to a file, close it, then return the current filename
''Function Declares
'Dim strConcatenated As String
'Dim lngFileNumber As Long
'Dim lngSectionCounter As Long
'Dim lngKeyCounter As Long
'    'Error Trap
'    On Error Resume Next
'    strConcatenated = ""
'    For lngSectionCounter = 1 To UBound(secFileSections, 1)
'        With secFileSections(lngSectionCounter)
'            'Add [SECTION]
'            strConcatenated = strConcatenated & "[" & .Name & "]" & vbCrLf
'            For lngKeyCounter = 1 To .KeyCount
'                'Add KEY=VALUE
'                strConcatenated = strConcatenated & .KeyValue(.IndexKey(lngKeyCounter)) & vbCrLf
'            Next
'        End With
'        'Add a blank line
'        strConcatenated = strConcatenated & vbCrLf
'    Next
'    lngFileNumber = FreeFile()
'    Open strFilename For Output As #lngFileNumber
'    Print #lngFileNumber, strConcatenated
'    Close #lngFileNumber
'    File = strFilename
'End Property

'Property Let File(ByVal strNewFileName As String)
''Created by Robin G Brown, 14/8/97
''Read from a file and create Sections and keys based on the file information
''Sub Declares
'Dim strFileLine As String
'Dim lngFileNumber As Long
'Dim lngAssignPosition As Long
'Dim strSectionPart As String
'Dim secCurrent As Section
'Dim strKeyPart As String
'Dim strValuePart As String
'Dim lngUbound As Long
'    'Error Trap
'    On Error Resume Next
'    strFilename = strNewFileName
'    lngFileNumber = FreeFile()
'    Open strFilename For Input As #lngFileNumber
'    Do While Not EOF(lngFileNumber)
'        Line Input #lngFileNumber, strFileLine
'        'Read line for data
'        If InStr(strFileLine, "[") > 0 Then
'            lngUbound = UBound(secFileSections, 1)
'            'If secCurrent is not nothing then set the previous section to be secCurrent
'            If Not secCurrent Is Nothing Then
'                Set secFileSections(lngUbound) = secCurrent
'            End If
'            'Create a new section in the array
'            If lngUbound = 0 Then
'                'Create a first section in the array
'                ReDim secFileSections(1 To lngUbound + 1)
'            Else
'                'Add a new section to the array
'                ReDim Preserve secFileSections(1 To lngUbound + 1)
'            End If
'            Set secCurrent = secFileSections(lngUbound + 1)
'            'Then create a new section
'            strSectionPart = Mid$(strFileLine, 2, Len(strFileLine) - 2)
'            Set secCurrent = New Section
'            secCurrent.Name = strSectionPart
'        ElseIf InStr(strFileLine, "=") > 0 Then
'            If Not secCurrent Is Nothing Then
'                'Only add keys if there is a current section
'                lngAssignPosition = InStr(strFileLine, "=")
'                strKeyPart = Left$(strFileLine, lngAssignPosition - 1)
'                strValuePart = Right$(strFileLine, Len(strFileLine) - lngAssignPosition)
'                'Add to the current section
'                secCurrent.KeyValue(strKeyPart) = strValuePart
'            End If
'        Else
'            'Not a valid line, so ignore
'        End If
'    Loop
'    'Add the last section, if there is one
'    lngUbound = UBound(secFileSections, 1)
'    If Not secCurrent Is Nothing Then
'        Set secFileSections(lngUbound) = secCurrent
'    End If
'    Close #lngFileNumber
'End Property

Property Get Text() As String
'Created by Robin G Brown, 18/8/97
'Output all secFileSections and keys to a Text string
'Function Declares
Dim strConcatenated As String
Dim lngFileNumber As Long
Dim lngSectionCounter As Long
Dim lngKeyCounter As Long
    'Error Trap
    On Error Resume Next
    strConcatenated = ""
    For lngSectionCounter = 1 To UBound(secFileSections, 1)
        With secFileSections(lngSectionCounter)
            'Add [SECTION]
            strConcatenated = strConcatenated & "[" & .Name & "]" & vbCrLf
            For lngKeyCounter = 1 To .KeyCount
                'Add KEY=VALUE
                strConcatenated = strConcatenated & .KeyValue(.IndexKey(lngKeyCounter)) & vbCrLf
            Next
        End With
        'Add a blank line
        strConcatenated = strConcatenated & vbCrLf
    Next
    Text = strConcatenated
End Property

Property Let Text(ByVal strScriptText As String)
'Created by Robin G Brown, 18/8/97
'Read from a text string and create secFileSections and keys based on the information
'Sub Declares
Dim strFileLine As String
Dim lngFileNumber As Long
Dim lngAssignPosition As Long
Dim strSectionPart As String
Dim secCurrent As Section
Dim strKeyPart As String
Dim strValuePart As String
Dim lngUbound As Long
Dim strTempScriptFile As String
    'Error Trap
    On Error Resume Next
    'Mega lazy cheat - save the text as a file, then read from that file, _
        this then requires no code rewrites! - RGB/18/8/97
    'This does set the strFileName property to the temp file though - RGB/10/9/97
    'Bit of a disk access overhead though!
    lngFileNumber = FreeFile()
    strTempScriptFile = App.Path & "\~SECTION.TXT"
    Open strTempScriptFile For Output As #lngFileNumber
    Print #lngFileNumber, strScriptText
    Close #lngFileNumber
    'Use the FileOpen method to reduce code duplication!
    Call Me.FileOpen(strTempScriptFile)
    Kill strTempScriptFile
End Property

Public Function FindSection(ByVal strSection As String) As Long
'Created by Robin G Brown, 15/8/97
'Find the index of the section with the given name
'Function Declares
Const conFunction = "FindSection"
Dim lngSectionCounter As Long
    'Error Trap
    On Error GoTo FindSection_ErrorHandler
    lngSectionCounter = 1
    Do While lngSectionCounter <= UBound(secFileSections, 1)
        If secFileSections(lngSectionCounter).Name = strSection Then
            FindSection = lngSectionCounter
            Exit Function
        End If
        lngSectionCounter = lngSectionCounter + 1
    Loop
    FindSection = -1
Exit Function
'Error Handler
FindSection_ErrorHandler:
    FindSection = -1
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conClass & ":" & conFunction, "Unexpected error : " & Err.Description
    End Select
    Exit Function
End Function

Public Sub AddSection(ByVal strSection As String, Optional ByVal varBefore As Variant)
'Created by Robin G Brown, 14/8/97
'Insert a new section into the array, in front of the given section, if necessary
'Sub Declares
Const conSub = "AddSection"
Const conAtEnd = -1
Dim lngInsertPosition As Long
Dim lngUbound As Long
Dim lngSectionCounter
    'Error Trap
    On Error Resume Next
    If IsMissing(varBefore) Then
        lngInsertPosition = conAtEnd
    Else
        If IsNumeric(varBefore) = True Then
            lngInsertPosition = CLng(varBefore)
        ElseIf VarType(varBefore) = vbString Then
            lngInsertPosition = FindSection(CStr(varBefore))
        Else
            'Invalid value
            Err.Raise vbObjectError, conClass & ":" & conSub, "invalid value for Before"
            Exit Sub
        End If
    End If
    Err.Clear
    lngUbound = UBound(secFileSections, 1)
    If lngInsertPosition > 0 And lngInsertPosition <= lngUbound Then
        'Insert at given position, moving all others back one position
        If lngUbound = 0 Then
            'Create a first section in the array
            ReDim secFileSections(1 To lngUbound + 1)
        Else
            'Add a new section to the array
            ReDim Preserve secFileSections(1 To lngUbound + 1)
        End If
        For lngSectionCounter = lngUbound + 1 To lngInsertPosition Step -1
            Set secFileSections(lngSectionCounter) = secFileSections(lngSectionCounter - 1)
        Next
        Set secFileSections(lngInsertPosition - 1) = New Section
        secFileSections(lngInsertPosition - 1).Name = strSection
    Else
        'Add to end
        If lngUbound = 0 Then
            'Create a first section in the array
            ReDim secFileSections(1 To lngUbound + 1)
        Else
            'Add a new section to the array
            ReDim Preserve secFileSections(1 To lngUbound + 1)
        End If
        Set secFileSections(lngUbound + 1) = New Section
        secFileSections(lngUbound + 1).Name = strSection
    End If
    If Err.Number <> 0 Then
        Err.Raise Err.Number, conClass & ":" & conSub, "Unexpected error : " & Err.Description
    End If
End Sub

Public Sub RemoveSection(ByVal strSection As String)
'Created by Robin G Brown, 14/8/97
'Remove a section from the collection
'Sub Declares
Const conSub = "RemoveSection"
Dim lngSectionIndex As Long
Dim lngSectionCounter As Long
Dim lngUbound As Long
    'Error Trap
    On Error Resume Next
    lngSectionIndex = FindSection(strSection)
    If lngSectionIndex > 0 Then
        'Delete this section by shuffling all the others back one place _
            and deleting the last one
        lngUbound = UBound(secFileSections, 1)
        For lngSectionCounter = lngSectionIndex To lngUbound - 1
            'Make sure the current object is destroyed first
            Set secFileSections(lngSectionCounter) = Nothing
            Set secFileSections(lngSectionCounter) = secFileSections(lngSectionCounter + 1)
        Next
        'Clear the last reference
        Set secFileSections(lngUbound) = Nothing
        'Remove the last element of the array
        ReDim Preserve secFileSections(1 To lngUbound - 1)
    Else
        Err.Raise vbObjectError, conClass & ":" & conSub, "Invalid Section."
    End If
End Sub

Public Function GetSectionSetting(ByVal strSection As String, ByVal strKey As String) As String
'Created by Robin G Brown, 14/8/97
'Return a value from the given section and key
'Function Declares
Const conFunction = "GetSectionSetting"
Dim lngSectionIndex As Long
    'Error Trap
    On Error GoTo GetSectionSetting_ErrorHandler
    lngSectionIndex = FindSection(strSection)
    If lngSectionIndex > 0 Then
        GetSectionSetting = secFileSections(lngSectionIndex).KeyValue(strKey)
    Else
        GetSectionSetting = ""
        Err.Raise vbObjectError, conClass & ":" & conFunction, "Invalid Section."
    End If
Exit Function
'Error Handler
GetSectionSetting_ErrorHandler:
    GetSectionSetting = ""
    Err.Raise Err.Number, conClass & ":" & conFunction, "Unexpected error : " & Err.Description
    Exit Function
End Function

Public Sub SetSectionSetting(ByVal strSection As String, ByVal strKey As String, ByVal strValue As String)
'Created by Robin G Brown, 14/8/97
'Set a value in the given section and key, _
    Section object will create a section, key and value _
    if it does not already exist
'Function Declares
Const conFunction = "SetSectionSetting"
Dim lngSectionIndex As Long
Dim lngKeyIndex As Long
Dim lngUbound As Long
    'Error Trap
    On Error Resume Next
    lngSectionIndex = FindSection(strSection)
    If lngSectionIndex > 0 Then
        secFileSections(lngSectionIndex).KeyValue(strKey) = strValue
    Else
        'Create a section
        lngUbound = UBound(secFileSections, 1)
        If lngUbound = 0 Then
            'Create a first section in the array
            ReDim secFileSections(1 To lngUbound + 1)
        Else
            'Add a new section to the array
            ReDim Preserve secFileSections(1 To lngUbound + 1)
        End If
        Set secFileSections(lngUbound + 1) = New Section
        secFileSections(lngUbound + 1).Name = strSection
        secFileSections(lngUbound + 1).KeyValue(strKey) = strValue
    End If
Exit Sub
    Err.Raise Err.Number, conClass & ":" & conFunction, "Unexpected error : " & Err.Description
    Exit Sub
End Sub

Private Sub Class_Terminate()
'Created by Robin G Brown, 14/8/97
'Sub Declares
    'Error Trap
    On Error Resume Next
    ReDim FilesecFileSections(0 To 0)
End Sub


