VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSVLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Created by Robin G Brown, 5th September 1997
'-----------------------------------------------------------------------------
'   This class module contains code for _
    a single CSV record line
'   Use crl as a prefix for this class
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---Implements Classes
'Implements CLASS_NAME_HERE
'---All Declares
Private Const conClass = "CSVLine"
'Public properties
'Private properties
Private lngCounter As Long
Private lngReturn As Long
'Replace Fields with name of Fields
Private strFields() As String

Private Sub Class_Initialize()
'Created by Robin G Brown, 5/9/97
'DESCRIPTION_OF_SUBROUTINE_FUNCTIONALITY_HERE
'Sub Declares
    'Error Trap
    On Error Resume Next
    'CODE_HERE
    strFields = ""
    Set FieldsOBJECT = Nothing
    If Err.Number <> 0 Then
        'ERROR_HANDLING_CODE_HERE
        Err.Raise Err.Number, conClass & ":Initialize", "Unexpected error : " & Err.Description
        Err.Clear
    End If
End Sub

Property Get Fields(ByVal Index As Variant) As String
'Created by Robin G Brown, 5/9/97
'DESCRIPTION_OF_FUNCTIONALITY_HERE
'Fields Declares
Const conFields = "Fields"
'Error Trap
    On Error Resume Next
    'CODE_HERE
    Fields = strFields
End Property

Property Let Fields(ByVal Index As Variant, ByVal strNewFields As String)
'Created by Robin G Brown, 5/9/97
'DESCRIPTION_OF_FUNCTIONALITY_HERE
'Sub Declares
Const conFields = "Fields"
    'Error Trap
    On Error Resume Next
    'CODE_HERE
    strFields = strNewFields
    If Err.Number <> 0 Then
        'ERROR_HANDLING_CODE_HERE
        Err.Raise Err.Number, conClass & ":" & conFields, "Unexpected error : " & Err.Description
    End If
End Property

Private Sub Class_Terminate()
'Created by Robin G Brown, 5/9/97
'DESCRIPTION_OF_SUBROUTINE_FUNCTIONALITY_HERE
'Sub Declares
    'Error Trap
    On Error Resume Next
    'CODE_HERE
    Set FieldsOBJECT = Nothing
    If Err.Number <> 0 Then
        'ERROR_HANDLING_CODE_HERE
        Err.Clear
    End If
End Sub


