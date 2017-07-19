VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Created by Robin G Brown, 20th May 1997
'-----------------------------------------------------------------------------
'   This class module contains code for _
    a simple stack class. Copied from a free web source code sample
'   Use stk as a prefix for this class
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "clsStack"
'Public properties
'Private properties
Private varData As New Collection
Private Const conFirstItem = 1

Private Sub Class_Initialize()
'Created by Robin G Brown, 20/5/97
'Default behaviour
'Sub Declares
    'Error Trap
    On Error Resume Next
    Set varData = New Collection
End Sub

Private Sub Class_Terminate()
'Created by Robin G Brown, 20/5/97
'Default behaviour
'Sub Declares
    'Error Trap
    On Error Resume Next
    Set varData = Nothing
End Sub

Public Sub Clear()
'Modified by Robin G Brown, 20/5/97
'Clear the stack
'Sub Declares
    'Error Trap
    On Error Resume Next
    Set varData = New Collection
End Sub

Property Get Count() As Long
'Modified by Robin G Brown, 20/5/97
'Returns the number of items in the stack
'Sub Declares
    'Error Trap
    On Error Resume Next
    Count = varData.Count
End Property

Public Function Peek() As Variant
'Modified by Robin G Brown, 20/5/97
'Returns the value of the next item on the stack without removing the item from the stack
'Sub Declares
    'Error Trap
    On Error Resume Next
    If varData.Count = 0 Then
        Peek = Null
    Else
        If VarType(varData.Item(conFirstItem)) = vbObject Then
            'If this is a stack of objects return the object
            Set Peek = varData.Item(conFirstItem)
        Else
            'Otherwise return the first item
            Peek = varData.Item(conFirstItem)
        End If
    End If
End Function

Public Function Pop() As Variant
'Modified by Robin G Brown, 20/5/97
'Removes and returns the next item from the stack.
'Sub Declares
    'Error Trap
    On Error Resume Next
    If varData.Count = 0 Then
        'No items so return null
        Pop = Null
    Else
        If VarType(varData(conFirstItem)) = vbObject Then
            'If this is a stack of objects return the object
            Set Pop = varData.Item(conFirstItem)
        Else
            'Otherwise return the first item
            Pop = varData.Item(conFirstItem)
        End If
        varData.Remove conFirstItem
    End If
End Function

Public Sub Push(ByVal varItem As Variant)
'Modified by Robin G Brown, 20/5/97
'Adds an item to the Stack
'Sub Declares
    'Error Trap
    On Error Resume Next
    If varData.Count = 0 Then
        'There are no other items so just add the new item
        varData.Add varItem
    Else
        'Add a new item in front of the current first item
        varData.Add varItem, , conFirstItem
    End If
End Sub
