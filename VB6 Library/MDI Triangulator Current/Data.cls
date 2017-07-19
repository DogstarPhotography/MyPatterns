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
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "BoilerPlateClass"
'Public properties
'Private properties
Private intCounter As Integer
Private intReturn As Integer
Private booProperty As Boolean
Private Sub Class_Initialize()
'Created by Robin G Brown, INSERT_DATE
'DESCRIPTION_OF_SUBROUTINE_FUNCTIONALITY_HERE
'Sub Declares
    'Error Trap
    On Error Resume Next
    'CODE_HERE
    If Err.Number <> 0 Then
        'ERROR_HANDLING_CODE_HERE
        Err.Clear
    End If
End Sub
Private Sub Class_Terminate()
'Created by Robin G Brown, INSERT_DATE
'DESCRIPTION_OF_SUBROUTINE_FUNCTIONALITY_HERE
'Sub Declares
    'Error Trap
    On Error Resume Next
    'CODE_HERE
    If Err.Number <> 0 Then
        'ERROR_HANDLING_CODE_HERE
        Err.Clear
    End If
End Sub
Property Get DelayedErrorPropertyGet() As Boolean
'Created by Robin G Brown, INSERT_DATE
'DESCRIPTION_OF_FUNCTIONALITY_HERE
'Function Declares
    'Error Trap
    On Error Resume Next
    DelayedErrorPropertyGet = False
    'CODE_HERE
    DelayedErrorPropertyGet = booProperty
    If Err.Number <> 0 Then
        'ERROR_HANDLING_CODE_HERE
        Err.Clear
    End If
    DelayedErrorPropertyGet = True
End Property
Property Let DelayedErrorPropertyLet(ByVal booNewProperty As Boolean)
'Created by Robin G Brown, INSERT_DATE
'DESCRIPTION_OF_FUNCTIONALITY_HERE
'Sub Declares
    'Error Trap
    On Error Resume Next
    'CODE_HERE
    booProperty = booNewProperty
    If Err.Number <> 0 Then
        'ERROR_HANDLING_CODE_HERE
        Err.Clear
    End If
End Property
Property Set DelayedErrorPropertySet(ByRef BPCBoilerPlateClass As BoilerPlateClass)
'Created by Robin G Brown, INSERT_DATE
'DESCRIPTION_OF_FUNCTIONALITY_HERE
'Sub Declares
    'Error Trap
    On Error Resume Next
    'CODE_HERE
    Set BPCBoilerPlateClass = Nothing
    If Err.Number <> 0 Then
        'ERROR_HANDLING_CODE_HERE
        Err.Clear
    End If
End Property
