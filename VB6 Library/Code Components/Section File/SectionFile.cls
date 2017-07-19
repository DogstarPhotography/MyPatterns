VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SectionFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Created by Robin G Brown, 6th August 1997
'-----------------------------------------------------------------------------
'   This class module contains code for _
    using a Section style file
'   Use sff as a prefix for this class
'-----------------------------------------------------------------------------
'   So what's a Section style file and how does it fit into strFileText
'
'   strFileText(1) = "[SECTION]"
'   strFileText(2) = "Key=Value"
'   strFileText(3) = "OtherKey=Value"
'   strFileText(4) = etc.
'
'   The lines must be seperated by vbCrLf 'characters' (actually a pair of characters)
'   which are stripped off as the file goes into the array
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "Editor"
'Public properties
'Private properties
Private lngArraycounter As Long
Private lngCurrentElement As Long
Private lngReturn As Long
Private strFileText() As String
Private strFilename As String

Private Sub Class_Initialize()
'Created by Robin G Brown, 6/8/97
'Sub Declares
    'Error Trap
    On Error Resume Next
    ReDim strFileText(1 To 1)
    strFileText(1) = ""
End Sub

Public Sub AddSection(ByVal strSection As String)
'Created by Robin G Brown, 7/8/97
'Add a new section
'Sub Declares
Const conSub = "AddSection"
Dim booNoKeyFlag As Boolean
    'Error Trap
    On Error GoTo AddSection_ErrorHandler
    'First make sure there is no section already
    For lngArraycounter = 1 To UBound(strFileText, 1)
        If InStr(strFileText(lngArraycounter), "[" & strSection & "]") > 0 Then
            'There is already a section so do not add one
            Exit Sub
        End If
    Next
    'Add an element into the array after the last element
    Call InsertElement(UBound(strFileText, 1))
    lngCurrentElement = UBound(strFileText, 1)
    'Set the element to strValue
    strFileText(lngCurrentElement) = "[" & strSection & "]"
Exit Sub
'Error Handler
AddSection_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conModule & ":" & conSub, "Unexpected error : " & Err.DeFileion
    End Select
    Exit Sub
End Sub

Public Sub AddKey(ByVal strSection As String, ByVal strKey As String, ByVal strValue As String)
'Created by Robin G Brown, 7/8/97
'Go through strFileText and set strValue to strKey within strSection, adding a key if required
'Sub Declares
Const conSub = "AddKey"
Dim booNoKeyFlag As Boolean
Dim booNoSectionFlag As Boolean
    'Error Trap
    On Error GoTo AddKey_ErrorHandler
    'Find the appropriate Section
    booNoSectionFlag = True
    For lngArraycounter = 1 To UBound(strFileText, 1)
        If InStr(strFileText(lngArraycounter), "[" & strSection & "]") > 0 Then
            lngCurrentElement = lngArraycounter
            booNoSectionFlag = False
            Exit For
        End If
    Next
    If booNoSectionFlag = True Then
        'If there is no section then exit
        Exit Sub
    End If
    'Find the Key bit
    booNoKeyFlag = True
    For lngArraycounter = 1 To UBound(strFileText, 1)
        If InStr(strFileText(lngArraycounter), strKey & "=") > 0 Then
            lngCurrentElement = lngArraycounter
            booNoKeyFlag = False
            Exit For
        ElseIf InStr(strFileText(lngArraycounter), "[") > 0 And InStr(strFileText(lngArraycounter), "[") > 0 Then
            lngCurrentElement = lngArraycounter
            Exit For
        End If
    Next
    'If there is no Key bit then create one
    If booNoKeyFlag = True Then
        'Add an element into the array after lngCurrentElement
        Call InsertElement(lngCurrentElement)
        lngCurrentElement = lngCurrentElement + 1
    End If
    'Set the element to strValue
    strFileText(lngCurrentElement) = strKey & "=" & strValue
Exit Sub
'Error Handler
AddKey_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conModule & ":" & conSub, "Unexpected error : " & Err.DeFileion
    End Select
    Exit Sub
End Sub

Public Sub SetKey(ByVal strSection As String, ByVal strKey As String, ByVal strValue As String)
'Created by Robin G Brown, 6/8/97
'Go through strFileText and set strValue to strKey within strSection, do not add anything
'Sub Declares
Const conSub = "SetKey"
Dim booNoKeyFlag As Boolean
Dim booNoSectionFlag As Boolean
    'Error Trap
    On Error GoTo SetKey_ErrorHandler
    'Find the appropriate Section
    booNoSectionFlag = True
    For lngArraycounter = 1 To UBound(strFileText, 1)
        If InStr(strFileText(lngArraycounter), "[" & strSection & "]") > 0 Then
            lngCurrentElement = lngArraycounter
            booNoSectionFlag = False
            Exit For
        End If
    Next
    If booNoSectionFlag = True Then
        'If there is no section then exit
        Exit Sub
    End If
    'Find the Key bit
    booNoKeyFlag = True
    For lngArraycounter = 1 To UBound(strFileText, 1)
        If InStr(strFileText(lngArraycounter), strKey & "=") > 0 Then
            lngCurrentElement = lngArraycounter
            booNoKeyFlag = False
            Exit For
        ElseIf InStr(strFileText(lngArraycounter), "[") > 0 And InStr(strFileText(lngArraycounter), "[") > 0 Then
            lngCurrentElement = lngArraycounter
            Exit For
        End If
    Next
    'If there is no Key bit then exit
    If booNoKeyFlag = True Then
        Exit Sub
    End If
    'Set the element to strValue
    strFileText(lngCurrentElement) = strKey & "=" & strValue
Exit Sub
'Error Handler
SetKey_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conModule & ":" & conSub, "Unexpected error : " & Err.DeFileion
    End Select
    Exit Sub
End Sub

Public Sub InsertElement(ByVal lngAfter As Long)
'Created by Robin G Brown, 6/8/97
'Insert an Element into an array after the element at position lngAfter
'Sub Declares
Const conSub = "InsertElement"
Dim lngBackwardCounter As Long
    'Error Trap
    On Error GoTo InsertElement_ErrorHandler
    ReDim Preserve strFileText(1 To UBound(strFileText, 1) + 1)
    'From the (new)last element, _
        go backwards copying the nth element into the nth+1 element, _
        being careful to stop before we go too far (lngAfter + 2)
    For lngBackwardCounter = UBound(strFileText, 1) To lngAfter + 2 Step -1
        strFileText(lngBackwardCounter) = strFileText(lngBackwardCounter - 1)
    Next
    strFileText(lngAfter + 1) = ""
Exit Sub
'Error Handler
InsertElement_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conModule & ":" & conSub, "Unexpected error : " & Err.Description
    End Select
    Exit Sub
End Sub

Property Get Text() As String
'Created by Robin G Brown, 6/8/97
'Output strFileText as a string
'Function Declares
Dim strConcatenated As String
    'Error Trap
    On Error Resume Next
    strConcatenated = ""
    For lngArraycounter = 1 To UBound(strFileText, 1)
        strConcatenated = strConcatenated & strFileText(lngArraycounter) & vbCrLf
    Next
    Text = strConcatenated
End Property

Property Let Text(ByVal strNewFile As String)
'Created by Robin G Brown, 6/8/97
'Read strNewFile and convert to an array
'   Spoiler Information
'   ## represents a vbCrLf pair
'   String
'   AB##CD##FG
'   1234567890
'     x   y  z
'   z = Len(String)
'   x = Instr(String, vbCrLf)
'   y = Instr(x + 2, String, vbCrLf)
'   Len("CD") = y - x - 2
'   "AB" = Left$(String, x - 1)
'   "FG" = Right$(String, z - y - 1)
'   "CD" = Mid$(String, x + 2, y - x - 2)
'Sub Declares
Const conProperty = "Text"
Dim lngXCrLfPos As Long
Dim lngYCrLfPos As Long
Dim lngZLen As Long
    'Error Trap
    On Error GoTo Text_ErrorHandler
    lngXCrLfPos = InStr(strNewFile, vbCrLf)
    lngZLen = Len(strNewFile)
    If lngXCrLfPos = 0 Then
        'Assume there's only one word
        ReDim strFileText(1 To 1)
        strFileText(1) = Trim$(strNewFile)
        Exit Property
    End If
    lngArraycounter = 1
    lngXCrLfPos = 0
    Do
        lngYCrLfPos = InStr(lngXCrLfPos + 2, strNewFile, vbCrLf)
        ReDim Preserve strFileText(1 To lngArraycounter)
        If lngYCrLfPos = 0 Then
            'Last word
            strFileText(lngArraycounter) = Trim$(StripCRLF(Right$(strNewFile, lngZLen - lngXCrLfPos - 1)))
            Exit Do
        Else
            'Middle word
            strFileText(lngArraycounter) = Trim$(StripCRLF(Mid$(strNewFile, lngXCrLfPos + 2, lngYCrLfPos - lngXCrLfPos - 2)))
            lngArraycounter = lngArraycounter + 1
        End If
        lngXCrLfPos = lngYCrLfPos
    Loop While lngXCrLfPos > 0
Exit Property
'Error Handler
Text_ErrorHandler:
    ReDim strFileText(1 To 1)
    Err.Raise Err.Number, conModule & ":" & conProperty, "Unexpected error : " & Err.Description
    Exit Property
End Property

Property Get File() As String
'Created by Robin G Brown, 7/8/97
'Output strFileText to a file and return the filename
'Function Declares
Dim strConcatenated As String
Dim lngFileNumber As Long
    'Error Trap
    On Error Resume Next
    strConcatenated = ""
    For lngArraycounter = 1 To UBound(strFileText, 1)
        strConcatenated = strConcatenated & strFileText(lngArraycounter) & vbCrLf
    Next
    lngFileNumber = FreeFile()
    Open strFilename For Output As #lngFileNumber
    Print #lngFileNumber, strConcatenated
    Close #lngFileNumber
    Text = strFilename
End Property

Property Let File(ByVal strNewFilename As String)
'Created by Robin G Brown, 7/8/97
'Read from a file into strFileText
'Sub Declares
Dim strFileLine As String
Dim lngFileNumber As Long
    'Error Trap
    On Error Resume Next
    strFilename = strNewFilename
    lngFileNumber = FreeFile()
    Open strFilename For Input As #lngFileNumber
    lngArraycounter = 1
    Do While Not EOF(lngFileNumber)
        Line Input #lngFileNumber, strFileLine
        ReDim Preserve strFileText(1 To lngArraycounter)
        strFileText(lngArraycounter) = Trim$(StripCRLF(strFileLine))
        lngArraycounter = lngArraycounter + 1
    Loop
    Close #lngFileNumber
End Property

Property Get Lines() As String
'Created by Robin G Brown, 7/8/97
'Return the nuber of elements in the file
'Function Declares
    'Error Trap
    On Error Resume Next
    Lines = UBound(strFileText, 1)
End Property

Private Sub Class_Terminate()
'Created by Robin G Brown, 6/8/97
'Sub Declares
    'Error Trap
    On Error Resume Next
    ReDim strFileText(1 To 1)
End Sub


