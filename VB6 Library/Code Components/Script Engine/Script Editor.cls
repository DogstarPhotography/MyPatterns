VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Editor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Created by Robin G Brown, 6th August 1997
'-----------------------------------------------------------------------------
'   This class module contains code for _
    editing script text, it holds a script as an array of strings - 1 element per line
'   Use scd as a prefix for this class
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "Editor"
'Public properties
'Private properties
Private lngScriptCounter As Long
Private lngReturn As Long
Private lngArraycounter As Long
Private lngCurrentElement As Long
Private strFileText() As String

Private Sub Class_Initialize()
'Created by Robin G Brown, 6/8/97
'Sub Declares
    'Error Trap
    On Error Resume Next
End Sub

Public Function SetProcessingPeriod(ByVal strProcessingPeriod As String) As String
'Created by Robin G Brown, 6/8/97
'Call SetOption for a number of actions that use processing period
'Sub Declares
Const conSub = "SetProcessingPeriod"
    'Error Trap
    On Error GoTo SetProcessingPeriod_ErrorHandler
    Call AddMultiKey("EXPORT", strProcessingPeriod)
    Call AddMultiKey("PURGE", strProcessingPeriod)
Exit Function
'Error Handler
SetProcessingPeriod_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conModule & ":" & conSub, "Unexpected error : " & Err.Description
    End Select
    Exit Function
End Function

Public Sub AddMultiKey(ByVal strSection As String, ByVal strProcessingPeriod As String)
'Created by Robin G Brown, 7/8/97
'Go through strFileText and set strProcessingPeriod to strKey within strSection, _
    strSection may be repeated. Add keys as required
'Sub Declares
Const conSub = "AddMultiKey"
Dim booNoKeyFlag As Boolean
Dim booExtendKeyFlag As Boolean
Dim intAssignPosition As Integer
    'Error Trap
    On Error GoTo AddMultiKey_ErrorHandler
    lngArraycounter = 1
    Do While lngArraycounter <= UBound(strFileText, 1)
        If InStr(strFileText(lngArraycounter), "[" & strSection & "]") > 0 Then
            'Found a section
            lngCurrentElement = lngArraycounter
            booNoKeyFlag = True
            Do While lngArraycounter <= UBound(strFileText, 1)
                If InStr(strFileText(lngArraycounter), "Options=") > 0 Then
                    'Found a key
                    lngCurrentElement = lngArraycounter
                    'Check for a processing period
                    If InStr(strFileText(lngCurrentElement), "ProcessingPeriod=") > 0 Then
                        'Replace the processing period using MId, the processing period MUST be six chars long
                        intAssignPosition = InStr(strFileText(lngCurrentElement), "ProcessingPeriod=") + 16
                        Mid(strFileText(lngCurrentElement), intAssignPosition, 6) = strProcessingPeriod
                    Else
                        'Extend the element by adding strProcessingPeriod
                        strFileText(lngCurrentElement) = strFileText(lngCurrentElement) & strProcessingPeriod
                    End If
                    booNoKeyFlag = False
                    Exit Do
                End If
                lngArraycounter = lngArraycounter + 1
            Loop
            If booNoKeyFlag = True Then
                'Add an element into the array after lngCurrentElement
                Call InsertElement(lngCurrentElement)
                lngCurrentElement = lngCurrentElement + 1
                lngArraycounter = lngCurrentElement
                'Set the element to strProcessingPeriod
                strFileText(lngCurrentElement) = "Options=ProcessingPeriod=" & strProcessingPeriod
            End If
        End If
        lngArraycounter = lngArraycounter + 1
    Loop
Exit Sub
'Error Handler
AddMultiKey_ErrorHandler:
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

Property Get Script() As String
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
    Script = strConcatenated
End Property

Property Let Script(ByVal strNewFile As String)
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
Const conProperty = "Script"
Dim lngXCrLfPos As Long
Dim lngYCrLfPos As Long
Dim lngZLen As Long
    'Error Trap
    On Error GoTo Script_ErrorHandler
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
Script_ErrorHandler:
    ReDim strFileText(1 To 1)
    Err.Raise Err.Number, conModule & ":" & conProperty, "Unexpected error : " & Err.Description
    Exit Property
End Property

Private Sub Class_Terminate()
'Created by Robin G Brown, 6/8/97
'Sub Declares
    'Error Trap
    On Error Resume Next
End Sub


