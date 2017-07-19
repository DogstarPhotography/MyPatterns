VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BoilerPlateClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Created by Robin G Brown, INSERT_LONG_DATE
'-----------------------------------------------------------------------------
'   This class module contains code for _
    DESCRIPTION_OF_CLASS_MODULE_FUNCTIONALITY_HERE
'   Use BPC as a prefix for this class
'-----------------------------------------------------------------------------
'   Special Instructions:
'   Replace PROPERTY with name of property
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---Implements Classes
'Implements CLASS_NAME_HERE
'---Events
'Instances of this class will need to be declared WithEvents if events are used
'EVENT_NAME_HERE is used to inform the object user _
    that EVENT_FUNCTIONALITY_HERE
'Public Event EVENT_NAME_HERE(ARGUMENTS_HERE)
'---Constants
Private Const conClass = "BoilerPlateClass"
'---Private Types
'---Variables
Private lngCounter As Long
Private lngReturn As Long
'---Public Properties
'---Private Properties
Private strPROPERTY As String
Private PROPERTYOBJECT As Object

Private Sub Class_Initialize()
'Created by Robin G Brown, INSERT_DATE
'DESCRIPTION_OF_SUBROUTINE_FUNCTIONALITY_HERE
'Sub Declares
    'Error Trap
    On Error Resume Next
    Err.Clear
    'CODE_HERE
    strPROPERTY = ""
    Set PROPERTYOBJECT = Nothing
    If Err.Number <> 0 Then
        'ERROR_HANDLING_CODE_HERE
        Err.Raise Err.Number, conClass & ":Initialize", "Unexpected error : " & Err.Description
    End If
End Sub

'-----------------------------------------------------------------------------
'   Event Handlers
'-----------------------------------------------------------------------------

'-----------------------------------------------------------------------------
'   Public Functions and Methods
'-----------------------------------------------------------------------------

'-----------------------------------------------------------------------------
'   Private Functions and Methods
'-----------------------------------------------------------------------------

Public Sub ImmediateErrorSub()
'Created by Robin G Brown, INSERT_DATE
'DESCRIPTION_OF_SUBROUTINE_FUNCTIONALITY_HERE
'Sub Declares
Const conSub = "ImmediateErrorSub"
    'Error Trap
    On Error GoTo ImmediateErrorSub_ErrorHandler
    'CODE_HERE
Exit Sub
'Error Handler
ImmediateErrorSub_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conClass & ":" & conSub, "Unexpected error : " & Err.Description
    End Select
    Exit Sub
End Sub

Public Function ImmediateErrorFunction() As String
'Created by Robin G Brown, INSERT_DATE
'DESCRIPTION_OF_FUNCTIONALITY_HERE
'Function Declares
Const conFunction = "ImmediateErrorFunction"
    'Error Trap
    On Error GoTo ImmediateErrorFunction_ErrorHandler
    'CODE_HERE
    ImmediateErrorFunction = ""
Exit Function
'Error Handler
ImmediateErrorFunction_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conClass & ":" & conFunction, "Unexpected error : " & Err.Description
    End Select
    ImmediateErrorFunction = ""
    Exit Function
End Function

'-----------------------------------------------------------------------------
'   Properties
'-----------------------------------------------------------------------------

Property Get PROPERTY() As String
'Created by Robin G Brown, INSERT_DATE
'DESCRIPTION_OF_FUNCTIONALITY_HERE
'Property Declares
Const conProperty = "PROPERTY"
'Error Trap
    On Error Resume Next
    'CODE_HERE
    PROPERTY = strPROPERTY
End Property

Property Let PROPERTY(ByVal strNewPROPERTY As String)
'Created by Robin G Brown, INSERT_DATE
'DESCRIPTION_OF_FUNCTIONALITY_HERE
'Property Declares
Const conProperty = "PROPERTY"
    'Error Trap
    On Error Resume Next
    'CODE_HERE
    strPROPERTY = strNewPROPERTY
    If Err.Number <> 0 Then
        'ERROR_HANDLING_CODE_HERE
        Err.Raise Err.Number, conClass & ":" & conProperty, "Unexpected error : " & Err.Description
    End If
End Property

Property Set PROPERTY(ByRef ObjectReference As Object)
'Created by Robin G Brown, INSERT_DATE
'DESCRIPTION_OF_FUNCTIONALITY_HERE
'Sub Declares
Const conProperty = "PROPERTY"
    'Error Trap
    On Error Resume Next
    'CODE_HERE
    Set PROPERTYOBJECT = ObjectReference
    If Err.Number <> 0 Then
        'ERROR_HANDLING_CODE_HERE
        Err.Raise Err.Number, conClass & ":" & conProperty, "Unexpected error : " & Err.Description
        Err.Clear
    End If
End Property

Private Sub Class_Terminate()
'Created by Robin G Brown, INSERT_DATE
'DESCRIPTION_OF_SUBROUTINE_FUNCTIONALITY_HERE
'Sub Declares
    'Error Trap
    On Error Resume Next
    Err.Clear
    'CODE_HERE
    Set PROPERTYOBJECT = Nothing
    If Err.Number <> 0 Then
        'ERROR_HANDLING_CODE_HERE
    End If
End Sub


