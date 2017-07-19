VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Script"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Created by Robin G Brown, 15th august 1997
'-----------------------------------------------------------------------------
'   This class module contains code for _
    using an Action style script file, created out of several _NON_UNIQUE_ Section objects, _
    held in an array
'   This class is a modified copy of 'File secScriptActions.cls' made on 15th August 1997 - RGB/15/8/97
'   Use scr as a prefix for this class
'-----------------------------------------------------------------------------
'   So what's a Action style file and how does it fit into strFileText
'
'   strFileText(1) = "[Action]"
'   strFileText(2) = "Key=Value"
'   strFileText(3) = "OtherKey=Value"
'   strFileText(4) = etc.
'
'   The lines must be seperated by vbCrLf 'characters' (actually a pair of characters) _
    which are stripped off as the file goes into the array
'
'   For scripts "[Action]" may be repeated more than once!
'   So there is a current position indicator which can be used _
    to implement sequential operations - Move/Find methods!
'
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conClass = "Script"
'Private properties
Private lngCounter As Long
Private lngReturn As Long
Private secScriptActions() As Section
Private lngCurrentActionIndex As Long
Private strFilename As String

Private Sub Class_Initialize()
'Created by Robin G Brown, 14/8/97
'Sub Declares
    'Error Trap
    On Error Resume Next
    ReDim secScriptActions(0 To 0)
    lngCurrentActionIndex = 0
    If Err.Number <> 0 Then
        Err.Raise Err.Number, conClass & ":Initialize", "Unexpected error : " & Err.Description
    End If
End Sub

Property Get LastAction() As Boolean
'Created by Robin G Brown, 18/8/97
'Return true if the current action is the last action
'Property Declares
Const conProperty = "LastAction"
'Error Trap
    On Error Resume Next
    If lngCurrentActionIndex = UBound(secScriptActions, 1) Then
        LastAction = True
    Else
        LastAction = False
    End If
End Property

Property Get Action(ByVal lngIndex As Long) As Section
'Created by Robin G Brown, 18/8/97
'Return the given action
'Property Declares
Const conProperty = "Action"
'Error Trap
    On Error Resume Next
    If lngIndex > 0 And lngIndex <= UBound(secScriptActions, 1) Then
        Set Action = secScriptActions(lngIndex)
        lngCurrentActionIndex = lngIndex
    Else
        Err.Raise vbObjectError, conClass & ":" & conProperty, "Invalid value for index."
    End If
End Property

Property Get ActionCount() As String
'Created by Robin G Brown, 18/8/97
'Return the number of actions there are in this object
'Property Declares
Const conProperty = "ActionCount"
'Error Trap
    On Error Resume Next
    ActionCount = UBound(secScriptActions, 1)
End Property

Property Get IndexAction(ByVal lngIndex As Long) As String
'Created by Robin G Brown, 18/8/97
'Return the Action name for a given index
'Sub Declares
Const conProperty = "IndexAction"
    'Error Trap
    On Error Resume Next
    IndexAction = secScriptActions(lngIndex).Name
End Property

Property Get ActionIndex() As Long
'Created by Robin G Brown, 18/8/97
'Return the index of the current action
'Property Declares
Const conProperty = "ActionIndex"
'Error Trap
    On Error Resume Next
    ActionIndex = lngCurrentActionIndex
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
    ReDim secScriptActions(0 To 0)
    lngCurrentActionIndex = 0
    'Now process the file
    Do While Not EOF(lngFileNumber)
        Line Input #lngFileNumber, strFileLine
        'Read line for data
        If InStr(strFileLine, "[") > 0 Then
            lngUbound = UBound(secScriptActions, 1)
            'If secCurrent is not nothing then set the previous section to be secCurrent
            If Not secCurrent Is Nothing Then
                Set secScriptActions(lngUbound) = secCurrent
            End If
            'Create a new section in the array
            If lngUbound = 0 Then
                ReDim secScriptActions(1 To lngUbound + 1)
            Else
                ReDim Preserve secScriptActions(1 To lngUbound + 1)
            End If
            Set secCurrent = secScriptActions(lngUbound + 1)
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
    lngUbound = UBound(secScriptActions, 1)
    If Not secCurrent Is Nothing Then
        Set secScriptActions(lngUbound) = secCurrent
    End If
    'Set the current action to the first action, if there is one
    If lngUbound > 0 Then
        lngCurrentActionIndex = 1
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
    For lngSectionCounter = 1 To UBound(secScriptActions, 1)
        With secScriptActions(lngSectionCounter)
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

Property Get Text() As String
'Created by Robin G Brown, 18/8/97
'Output all secScriptActions and keys to a Text string
'Function Declares
Dim strConcatenated As String
Dim lngFileNumber As Long
Dim lngActionCounter As Long
Dim lngKeyCounter As Long
    'Error Trap
    On Error Resume Next
    strConcatenated = ""
    For lngActionCounter = 1 To UBound(secScriptActions, 1)
        With secScriptActions(lngActionCounter)
            'Add [Action]
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
'Read from a text string and create secScriptActions and keys based on the information
'Sub Declares
Dim strFileLine As String
Dim lngFileNumber As Long
Dim lngAssignPosition As Long
Dim strActionPart As String
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
    strTempScriptFile = App.Path & "\~Action.TXT"
    Open strTempScriptFile For Output As #lngFileNumber
    Print #lngFileNumber, strScriptText
    Close #lngFileNumber
    'Use the FileOpen method to reduce code duplication!
    Call Me.FileOpen(strTempScriptFile)
    'Destroy the temporary file
    Kill strTempScriptFile
    'Set lngCurrentActionIndex to the first Action
    lngCurrentActionIndex = 1
End Property

Private Sub MoveFirstAction()
'Created by Robin G Brown, 15/8/97
'Move to the first action
'Sub Declares
Const conSub = "MoveFirstAction"
    'Error Trap
    On Error Resume Next
    If UBound(secScriptActions, 1) > 0 Then
        lngCurrentActionIndex = 1
    Else
        lngCurrentActionIndex = -1
        Err.Raise vbObjectError, conClass & ":" & conSub, "No actions in this file."
    End If
End Sub

Private Sub MoveNextAction()
'Created by Robin G Brown, 15/8/97
'Move to the first action
'Sub Declares
Const conSub = "MoveNextAction"
    'Error Trap
    On Error Resume Next
    If UBound(secScriptActions, 1) < lngCurrentActionIndex Then
        lngCurrentActionIndex = lngCurrentActionIndex + 1
    Else
        lngCurrentActionIndex = -1
        Err.Raise vbObjectError, conClass & ":" & conSub, "No more actions in this file."
    End If
End Sub


Private Function FindFirstAction(ByVal strAction As String) As Long
'Created by Robin G Brown, 15/8/97
'Find the index of the first Action with the given name
'Function Declares
Const conFunction = "FindFirstAction"
Dim lngActionCounter As Long
    'Error Trap
    On Error GoTo FindFirstAction_ErrorHandler
    lngActionCounter = 1
    Do While lngActionCounter <= UBound(secScriptActions, 1)
        If secScriptActions(lngActionCounter).Name = strAction Then
            FindFirstAction = lngActionCounter
            lngCurrentActionIndex = lngActionCounter
            Exit Function
        End If
        lngActionCounter = lngActionCounter + 1
    Loop
    FindFirstAction = -1
    lngCurrentActionIndex = -1
Exit Function
'Error Handler
FindFirstAction_ErrorHandler:
    FindFirstAction = -1
    lngCurrentActionIndex = -1
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conClass & ":" & conFunction, "Unexpected error : " & Err.Description
    End Select
    Exit Function
End Function

Private Function FindNextAction(ByVal strAction As String) As Long
'Created by Robin G Brown, 15/8/97
'Find the index of the next Action after secScriptActions(lngCurrentActionIndicator) with the given name _
    and set lngCurrentActionIndicator to that Action's index
'Function Declares
Const conFunction = "FindNextAction"
Dim lngActionCounter As Long
    'Error Trap
    On Error GoTo FindNextAction_ErrorHandler
    lngActionCounter = lngCurrentActionIndex
    Do While lngActionCounter <= UBound(secScriptActions, 1)
        If secScriptActions(lngActionCounter).Name = strAction Then
            FindNextAction = lngActionCounter
            lngCurrentActionIndex = lngActionCounter
            Exit Function
        End If
        lngActionCounter = lngActionCounter + 1
    Loop
    FindNextAction = -1
    lngCurrentActionIndex = -1
Exit Function
'Error Handler
FindNextAction_ErrorHandler:
    FindNextAction = -1
    lngCurrentActionIndex = -1
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conClass & ":" & conFunction, "Unexpected error : " & Err.Description
    End Select
    Exit Function
End Function

Public Sub AddAction(ByVal strAction As String, Optional ByVal varBefore As Variant)
'Created by Robin G Brown, 14/8/97
'Insert a new Action into the array and make it the current Action
'Sub Declares
Const conSub = "AddAction"
Const conAtEnd = -1
Dim lngInsertPosition As Long
Dim lngUbound As Long
Dim lngActionCounter
    'Error Trap
    On Error Resume Next
    If IsMissing(varBefore) Then
        lngInsertPosition = conAtEnd
    Else
        If IsNumeric(varBefore) = True Then
            lngInsertPosition = CLng(varBefore)
        'ElseIf VarType(varBefore) = vbString Then
        '    lngInsertPosition = FindAction(CStr(varBefore))
        Else
            'Insert before the current Action!
            lngInsertPosition = lngCurrentActionIndex
        End If
    End If
    Err.Clear
    lngUbound = UBound(secScriptActions, 1)
    If lngInsertPosition > 0 And lngInsertPosition <= lngUbound Then
        'Insert at given position, moving all others back one position
        If lngUbound = 0 Then
            'Create a first action in the array
            ReDim secScriptActions(1 To lngUbound + 1)
        Else
            'Add a new action to the array
            ReDim Preserve secScriptActions(1 To lngUbound + 1)
        End If
        For lngActionCounter = lngUbound + 1 To lngInsertPosition Step -1
            Set secScriptActions(lngActionCounter) = secScriptActions(lngActionCounter - 1)
        Next
        Set secScriptActions(lngInsertPosition - 1) = New Section
        secScriptActions(lngInsertPosition - 1).Name = strAction
        lngCurrentActionIndex = lngInsertPosition - 1
    Else
        'Add to end
        If lngUbound = 0 Then
            'Create a first action in the array
            ReDim secScriptActions(1 To lngUbound + 1)
        Else
            'Add a new action to the array
            ReDim Preserve secScriptActions(1 To lngUbound + 1)
        End If
        Set secScriptActions(lngUbound + 1) = New Section
        secScriptActions(lngUbound + 1).Name = strAction
        'Note that this makes secCurrentScriptAction the last Action!
        lngCurrentActionIndex = lngUbound + 1
    End If
    If Err.Number <> 0 Then
        Err.Raise Err.Number, conClass & ":" & conSub, "Unexpected error : " & Err.Description
    End If
End Sub

Public Sub RemoveAction()
'Created by Robin G Brown, 14/8/97
'Remove the current Action
'Sub Declares
Const conSub = "RemoveAction"
Dim lngActionIndex As Long
Dim lngActionCounter As Long
Dim lngUbound As Long
    'Error Trap
    On Error Resume Next
    lngActionIndex = lngCurrentActionIndex
    If lngActionIndex > 0 Then
        'Delete this Action by shuffling all the others back one place _
            and deleting the last one
        lngUbound = UBound(secScriptActions, 1)
        For lngActionCounter = lngActionIndex To lngUbound - 1
            'Make sure the current object is destroyed first
            Set secScriptActions(lngActionCounter) = Nothing
            Set secScriptActions(lngActionCounter) = secScriptActions(lngActionCounter + 1)
        Next
        'Clear the last reference
        Set secScriptActions(lngUbound) = Nothing
        'Remove the last element of the array
        ReDim Preserve secScriptActions(1 To lngUbound - 1)
    Else
        Err.Raise vbObjectError, conClass & ":" & conSub, "Invalid Action."
    End If
End Sub

Public Function GetKey(ByVal strKey As String) As String
'Created by Robin G Brown, 14/8/97
'Return a value from the current action and given key
'Function Declares
Const conFunction = "GetKey"
    'Error Trap
    On Error Resume Next
    GetKey = secScriptActions(lngCurrentActionIndex).KeyValue(strKey)
End Function

Public Sub SetKey(ByVal strKey As String, ByVal strValue As String)
'Created by Robin G Brown, 14/8/97
'Set a value in the current Action and given key, _
    Action object will create a key and value _
    if it does not already exist
'Function Declares
Const conFunction = "SetKey"
    'Error Trap
    On Error Resume Next
    secScriptActions(lngCurrentActionIndex).KeyValue(strKey) = strValue
End Sub

Property Get BookMark() As Long
'Created by Robin G Brown, 15/8/97
'Return the index of secCurrentScriptAction
'Property Declares
Const conProperty = "BookMark"
'Error Trap
    On Error Resume Next
    BookMark = lngCurrentActionIndex
End Property

Property Let BookMark(ByVal lngBookMark As Long)
'Created by Robin G Brown, 15/8/97
'Set secCurrentScriptAction to the given Action by using the bookmark as an index
'Sub Declares
Const conProperty = "BookMark"
    'Error Trap
    On Error Resume Next
    If lngBookMark > 0 And lngBookMark <= UBound(secScriptActions, 1) Then
        lngCurrentActionIndex = lngBookMark
    Else
        Err.Raise vbObjectError, conClass & ":" & conProperty, "Invalid Bookmark, not set."
    End If
End Property

Private Sub Class_Terminate()
'Created by Robin G Brown, 14/8/97
'Sub Declares
    'Error Trap
    On Error Resume Next
    ReDim secScriptActions(0 To 0)
End Sub


