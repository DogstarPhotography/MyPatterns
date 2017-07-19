VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Section"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Created by Robin G Brown, 14th August, 1997
'-----------------------------------------------------------------------------
'   This class module contains code for _
    a section object, which contains an array of unique keys, _
    a collection of these objects is used to create a section file
'   Use sec as a prefix for this class
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conClass = "Section"
'Private properties
Private strKeys() As String

Private Sub Class_Initialize()
'Created by Robin G Brown, 14/8/97
'Sub Declares
    'Error Trap
    On Error Resume Next
    ReDim strKeys(1 To 2, 0 To 0)
    If Err.Number <> 0 Then
        Err.Raise Err.Number, conClass & ":Initialize", "Unexpected error : " & Err.Description
    End If
End Sub

Property Get Index() As Long
'Created by Robin G Brown, 15/8/97
'Return the Index of this section, index has no internal function, it is used merely as an identifier
'Property Declares
Const conProperty = "Index"
'Error Trap
    On Error Resume Next
    Index = CLng(strKeys(2, 0))
End Property

Property Let Index(ByVal lngIndex As Long)
'Created by Robin G Brown, 15/8/97
'Set the Index of this section, index has no internal function, it is used merely as an identifier
'Sub Declares
Const conProperty = "Index"
    'Error Trap
    On Error Resume Next
    strKeys(2, 0) = CStr(lngIndex)
    If Err.Number <> 0 Then
        Err.Raise Err.Number, conClass & ":" & conProperty, "Unexpected error : " & Err.Description
    End If
End Property

Property Get Name() As String
'Created by Robin G Brown, 14/8/97
'Return the name of this section
'Property Declares
Const conProperty = "Name"
'Error Trap
    On Error Resume Next
    Name = strKeys(1, 0)
End Property

Property Let Name(ByVal strName As String)
'Created by Robin G Brown, 14/8/97
'Set the name of this section
'Sub Declares
Const conProperty = "Name"
    'Error Trap
    On Error Resume Next
    strKeys(1, 0) = strName
    If Err.Number <> 0 Then
        Err.Raise Err.Number, conClass & ":" & conProperty, "Unexpected error : " & Err.Description
    End If
End Property

Property Get KeyCount() As Long
'Created by Robin G Brown, 14/8/97
'Return the number of Keys
'Sub Declares
Const conProperty = "KeyCount"
Dim lngIndex As Long
    'Error Trap
    On Error Resume Next
    KeyCount = UBound(strKeys, 2)
End Property

Property Get IndexKey(ByVal lngIndex As Long) As String
'Created by Robin G Brown, 14/8/97
'Return the Key for a given index
'Sub Declares
Const conProperty = "IndexKey"
    'Error Trap
    On Error Resume Next
    IndexKey = strKeys(1, lngIndex)
End Property

Property Get KeyValue(ByVal strKey As String) As String
'Created by Robin G Brown, 14/8/97
'Return a Value for a key
'Sub Declares
Const conProperty = "Key"
Dim lngIndex As Long
    'Error Trap
    On Error Resume Next
    'Search for Key
    lngIndex = 1
    Do While lngIndex <= UBound(strKeys, 2)
        If strKeys(1, lngIndex) = strKey Then
        'Key found, change KeyValue, then exit
            KeyValue = strKeys(2, lngIndex)
            Exit Property
        End If
        lngIndex = lngIndex + 1
    Loop
    'Key not found, return an error
    Err.Raise vbObjectError, conClass & ":" & conProperty, "Key not found."
End Property

Property Let KeyValue(ByVal strKey As String, ByVal strValue As String)
'Created by Robin G Brown, 14/8/97
'Add or set a key and Value
'Sub Declares
Const conProperty = "KeyValue"
Dim lngNewKey As Long
Dim lngIndex As Long
    'Error Trap
    On Error Resume Next
    'Search for Key
    lngIndex = 1
    Do While lngIndex <= UBound(strKeys, 2)
        If strKeys(1, lngIndex) = strKey Then
            'Key found, change KeyValue, then exit
            strKeys(2, lngIndex) = strValue
            Exit Property
        End If
        lngIndex = lngIndex + 1
    Loop
    'Key not found, add key and KeyValue
    Err.Clear
    lngNewKey = UBound(strKeys, 2) + 1
    ReDim Preserve strKeys(1 To 2, 0 To lngNewKey)
    strKeys(1, lngNewKey) = strKey
    strKeys(2, lngNewKey) = strValue
    If Err.Number <> 0 Then
        Err.Raise Err.Number, conClass & ":" & conProperty, "Unexpected error : " & Err.Description
    End If
End Property

Public Sub RemoveKey(ByVal strKey As String)
'Created by Robin G Brown, 14/8/97
'Remove the given key
'Sub Declares
Const conSub = "RemoveKey"
Dim lngAssignPosition As Long
Dim lngIndex As Long
    'Error Trap
    On Error Resume Next
    'Search for Key
    lngIndex = 1
    Do While lngIndex <= UBound(strKeys, 2)
        If strKeys(1, lngIndex) = strKey Then
            'Key found, remove entry by shifting elements _
                and removing the last element, then exit
            Do While lngIndex < UBound(strKeys, 2)
                strKeys(1, lngIndex) = strKeys(1, lngIndex + 1)
                strKeys(2, lngIndex) = strKeys(2, lngIndex + 1)
                lngIndex = lngIndex + 1
            Loop
            ReDim Preserve strKeys(1 To 2, 0 To UBound(strKeys, 2) - 1)
            Exit Sub
        End If
        lngIndex = lngIndex + 1
    Loop
    'Key not found, return an error
    Err.Raise vbObjectError, conClass & ":" & conSub, "Key not found."
End Sub

Private Sub Class_Terminate()
'Created by Robin G Brown, 14/8/97
'Sub Declares
    'Error Trap
    On Error Resume Next
    ReDim strKeys(1 To 2, 0 To 0)
End Sub


