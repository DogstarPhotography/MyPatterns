VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Created by Robin G Brown, 6th May 1997
'-----------------------------------------------------------------------------
'   This class module contains code for _
    holding and manipulating the data for the grid
'   Use dtg as a prefix for this class
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
Private varData() As Variant
'Error Constants
Private Const errArray = 55001

Private Sub Class_Initialize()
'Created by Robin G Brown, 6/5/97
'Default Behaviour
'Sub Declares
    'Error Trap
    On Error Resume Next
    'Initialise the array
    ReDim varData(1 To 1)
    If Err.Number <> 0 Then
        'ERROR_HANDLING_CODE_HERE
        Err.Clear
    End If
End Sub

Private Sub Class_Terminate()
'Created by Robin G Brown, 6/5/97
'Default Behaviour
'Sub Declares
    'Error Trap
    On Error Resume Next
    'Nobble the array
    ReDim varData(1 To 1)
    If Err.Number <> 0 Then
        'ERROR_HANDLING_CODE_HERE
        Err.Clear
    End If
End Sub

Public Function LoadData(ByRef varArray() As Variant) As Boolean
'Created by Robin G Brown, 2/5/97
'Fill varData() from an array of data - varArray()
'Function Declares
Const conFunction = "LoadData"
Dim intLBoundD1 As Integer
Dim intUBoundD1 As Integer
Dim intDimension1Counter As Integer
Dim intLBoundD2 As Integer
Dim intUBoundD2 As Integer
Dim intDimension2Counter As Integer
    'Error Trap
    On Error GoTo LoadData_ErrorHandler
    'Read the size of the array
    intLBoundD1 = LBound(varArray, 1)
    intUBoundD1 = UBound(varArray, 1)
    intLBoundD2 = LBound(varArray, 1)
    intUBoundD2 = UBound(varArray, 1)
    'Now we know how big it is we can load the data into our own array
    ReDim varData(intLBoundD1 To intUBoundD1, intLBoundD2 To intUBoundD2)
    For intDimension1Counter = intLBoundD2 To intUBoundD1
        For intDimension2Counter = intLBoundD2 To intUBoundD2
            varData(intDimension1Counter, intDimension2Counter) = _
                varArray(intDimension1Counter, intDimension2Counter)
        Next
    Next
    LoadData = True
Exit Function
'Error Handler
LoadData_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conFunction
    End Select
    ReDim varData(1 To 1, 1 To 1)
    LoadData = False
    Exit Function
End Function

Public Sub CopyData(ByRef varArray() As Variant)
'Created by Robin G Brown, 2/5/97
'Copy varData into varArray
'Function Declares
Const conSub = "CopyData"
Dim intLBoundD1 As Integer
Dim intUBoundD1 As Integer
Dim intDimension1Counter As Integer
Dim intLBoundD2 As Integer
Dim intUBoundD2 As Integer
Dim intDimension2Counter As Integer
    'Error Trap
    On Error GoTo CopyData_ErrorHandler
    'Read the size of the array
    intLBoundD1 = LBound(varData, 1)
    intUBoundD1 = UBound(varData, 1)
    intLBoundD2 = LBound(varData, 2)
    intUBoundD2 = UBound(varData, 2)
    'Now we know how big it is we can load the data into our own array
    ReDim varArray(intLBoundD1 To intUBoundD1, intLBoundD2 To intUBoundD2)
    For intDimension1Counter = intLBoundD2 To intUBoundD1
        For intDimension2Counter = intLBoundD2 To intUBoundD2
            varArray(intDimension1Counter, intDimension2Counter) = _
                varData(intDimension1Counter, intDimension2Counter)
        Next
    Next
Exit Sub
'Error Handler
CopyData_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    ReDim varArray(1 To 1, 1 To 1)
    Exit Sub
End Sub

Property Get GetElement(ByVal intD1Pos As Integer, ByVal intD2Pos As Integer) As Variant
'Created by Robin G Brown, 2/5/97
'Return the value of the given element
'Function Declares
    'Error Trap
    On Error Resume Next
    GetElement = varData(intD1Pos, intD2Pos)
    If Err.Number <> 0 Then
        GetElement = Null
        Err.Clear
    End If
End Property

Property Let SetElement(ByVal intD1Pos As Integer, ByVal intD2Pos As Integer, ByVal varValue As Variant)
'Created by Robin G Brown, 2/5/97
'Set the value of the given element
'Sub Declares
    'Error Trap
    On Error Resume Next
    varData(intD1Pos, intD2Pos) = varValue
    If Err.Number <> 0 Then
        'Raise an error??
        'ThrowError errArray
    End If
End Property

Public Sub ThrowError(ByVal errNumber As Long)
'Created by Robin G Brown, 2/5/97
'Raise an error
'Sub Declares
    'Error Trap
    On Error Resume Next
    Select Case errNumber
    Case errArray
        Err.Raise errArray, conModule, "Array error"
    Case Else
        Err.Raise 55000, conModule, "Undefined Error"
    End Select
End Sub

Public Sub CreateData()
'Created by Robin G Brown, 2/5/97
'Fill varData() from an array of data - varArray()
'Function Declares
Const conFunction = "CreateData"
Dim intLBoundD1 As Integer
Dim intUBoundD1 As Integer
Dim intDimension1Counter As Integer
Dim intLBoundD2 As Integer
Dim intUBoundD2 As Integer
Dim intDimension2Counter As Integer
Dim sngRandom As Single
    'Error Trap
    On Error Resume Next
    'Read the size of the array
    intLBoundD1 = 1
    intUBoundD1 = 10
    intLBoundD2 = 1
    intUBoundD2 = 10
    'Now we know how big it is we can load the data into our own array
    ReDim varData(intLBoundD1 To intUBoundD1, intLBoundD2 To intUBoundD2)
    For intDimension1Counter = intLBoundD2 To intUBoundD1
        sngRandom = Rnd * 100
        For intDimension2Counter = intLBoundD2 To intUBoundD2
            varData(intDimension1Counter, intDimension2Counter) = Format$(sngRandom * intDimension2Counter, "000.00")
        Next
    Next
End Sub

