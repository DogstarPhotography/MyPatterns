Attribute VB_Name = "modTreeViewLib"
'Created by Robin G Brown, 11th April 1997
'-----------------------------------------------------------------------------
'   This module contains code for _
    handling treeview controls
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "modTreeViewLib"
Private Const conDefaultTreeViewImage = "default"
Private intCounter As Integer
Private intReturn As Integer
Private Const conCopyString = "_Copy"
Public Function GetDescendants(ByRef nodSource As Node) As Integer
'Created by Robin G Brown, 16/4/97
'Iteratively count the total number of Descendants of this ndoe
'Function Declares
Dim nodCurrent As Node
Dim nodChild As Node
Dim intDescendants As Integer
    'Error Trap
    On Error GoTo GetDescendants_ErrorHandler
    'We have not found any Descendants yet
    intDescendants = 0
    'Goto first child of source node
    Set nodCurrent = nodSource.Child
    'If it exists count it/goto next child
    Do While Not nodCurrent Is Nothing
        intDescendants = intDescendants + 1
        'Check for children of child of source node
        Set nodChild = nodCurrent.Child
        If Not nodCurrent.Child Is Nothing Then
            'If it has any children count them, by calling this function iteratively
            intDescendants = intDescendants + GetDescendants(nodCurrent)
        End If
        Set nodCurrent = nodCurrent.Next
    Loop
    GetDescendants = intDescendants
Exit Function
GetDescendants_ErrorHandler:
    GetDescendants = -1
    Exit Function
End Function
Public Function GetSiblings(ByRef nodSource As Node) As Integer
'Created by Robin G Brown, 16/4/97
'Returns the number of siblings a node has, _
    INCLUDING itself, _
    returns 0 on error
'Function Declares
Const conFunction = "GetSiblings"
    'Error Trap
    On Error GoTo GetSiblings_ErrorHandler
    GetSiblings = nodSource.Parent.Children 'Haha, that was easy! - RGB
Exit Function
'Error Handler
GetSiblings_ErrorHandler:
    GetSiblings = 0
    Exit Function
End Function
Public Function GetIndentation(ByRef nodSource As Node) As Integer
'Created by Robin G Brown, 16/4/97
'Finds out how deeply 'nested' a node is, _
    returns 0 if this is the 'root' node, _
    + 1 for each level of indentation. _
    Returns -1 on error
'Function Declares
Const conFunction = "GetIndentation"
Dim intIndent As Integer
Dim nodParent As Node
Dim nodCurrent As Node
    'Error Trap
    On Error GoTo GetIndentation_ErrorHandler
    intIndent = 0
    Set nodParent = nodSource.Parent
    If nodParent Is Nothing Then
        GetIndentation = 0
    Else
        Do
            intIndent = intIndent + 1
            Set nodCurrent = nodParent
            Set nodParent = nodCurrent.Parent
        Loop While Not nodParent Is Nothing
        GetIndentation = intIndent
    End If
Exit Function
'Error Handler
GetIndentation_ErrorHandler:
    GetIndentation = -1
    Exit Function
End Function
Public Sub CopyBranch(ByRef tvwControl As TreeView, ByRef nodSource As Node, ByVal strTargetKey As String)
'Created by Robin G Brown, 14/4/97
'Attach nodSource to tvwControl under strTargetKey
'Assumptions : _
    adding conCopyString (see above) to end of key will not cause problems _
    there is no requirement to remove nodSource, _
    nodSource does NOT come from tvwControl
'Sub Declares
Const conSub = "CopyBranch"
    'Error Trap
    On Error GoTo CopyBranch_ErrorHandler
    'Any setup requirements
    'No setup requirements!
    'Call CopyNode to copy branch iteratively
    Call CopyNode(tvwControl, nodSource, strTargetKey)
Exit Sub
'Error Handler
CopyBranch_ErrorHandler:
    MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    Exit Sub
End Sub
Public Sub CopyNode(ByRef tvwControl As TreeView, ByRef nodSource As Node, ByVal strTargetKey As String)
'Created by Robin G Brown, 14/11/97
'Call iteratively to copy nodes of a branch
'Sub Declares
Const conSub = "CopyNode"
Dim nodCurrent As Node
Dim nodParent As Node
Dim nodChild As Node
Dim nodtemp As Node
    'Error Trap
    On Error GoTo CopyNode_ErrorHandler
    'Copy first node
    Set nodParent = AddNode(tvwControl, nodSource, strTargetKey)
    'If it was visible before make it visible now
    If nodSource.Expanded = True Then
        nodParent.Expanded = True
    End If
    'If it was visible before make it visible now
    If nodSource.Visible = True Then
        nodParent.EnsureVisible
    End If
    'Goto first child of source node
    Set nodCurrent = nodSource.Child.FirstSibling
    'If it exists copy it/goto next child
    Do While Not nodCurrent Is Nothing
        'Check for children of child of source node
        Set nodChild = nodCurrent.Child
        If nodCurrent.Child Is Nothing Then
            'No children so just copy it
            Set nodtemp = AddNode(tvwControl, nodCurrent, nodParent.Key)
            'If it was visible before make it visible now
            If nodCurrent.Expanded = True Then
                nodtemp.Expanded = True
            End If
            'If it was visible before make it visible now
            If nodCurrent.Visible = True Then
                nodtemp.EnsureVisible
            End If
        Else
            'If it has any children copy them, by calling this function iteratively
            Call CopyNode(tvwControl, nodCurrent, nodParent.Key)
        End If
        Set nodCurrent = nodCurrent.Next
    Loop
Exit Sub
'Error Handler
CopyNode_ErrorHandler:
    Exit Sub
End Sub
Public Function AddNode(ByRef tvwControl As TreeView, ByRef nodSource As Node, ByVal strTargetKey As String) As Node
'Created by Robin G Brown, 14/4/97
'Copy nodSource to tvwControl under strTargetkey, _
    correcting key values to prevent duplication and crashes!
'Sub Declares
Const conSub = "AddNode"
    'Error Trap
    On Error GoTo AddNode_ErrorHandler
    If CheckForDuplicateKey(tvwControl, nodSource.Key) = True Then
        Set AddNode = tvwControl.Nodes.Add(strTargetKey, tvwChild, nodSource.Key & conCopyString, nodSource.Text, nodSource.Image, nodSource.SelectedImage)
        'The following line will show the key in the text, not the original text
        'Set AddNode = tvwControl.Nodes.Add(strTargetKey, tvwChild, nodSource.Key & conCopyString, nodSource.Key & conCopyString, nodSource.Image, nodSource.SelectedImage)
    Else
        Set AddNode = tvwControl.Nodes.Add(strTargetKey, tvwChild, nodSource.Key, nodSource.Text, nodSource.Image, nodSource.SelectedImage)
        'The following line will show the key in the text, not the original text
        'Set AddNode = tvwControl.Nodes.Add(strTargetKey, tvwChild, nodSource.Key, nodSource.Key, nodSource.Image, nodSource.SelectedImage)
    End If
Exit Function
'Error Handler
AddNode_ErrorHandler:
    Exit Function
End Function
Public Sub CorrectBranchKeys(ByRef tvwControl As TreeView, ByVal nodCheck As Node)
'Created by Robin G Brown, 14/4/97
'Correct all keys in this branch
'Sub Declares
Const conSub = "CorrectBranchKeys"
Dim nodChild As Node
Dim intNext As Integer
    'Error Trap
    On Error GoTo CorrectBranchKeys_ErrorHandler
    'Correct current node
    If CheckForDuplicateKey(tvwControl, nodCheck.Key) = True Then
        nodCheck.Key = nodCheck & "_Copy"
        nodCheck.Text = nodCheck.Key
    End If
    'Goto first child
    Set nodCheck = nodCheck.Child.FirstSibling
    'If it exists correct that child/carry on to any children
    Do While Not nodCheck Is Nothing
        'Check for children
        Set nodChild = nodCheck.Child
        If nodCheck.Child Is Nothing Then
            'Correct current node
            If CheckForDuplicateKey(tvwControl, nodCheck.Key) = True Then
                nodCheck.Key = nodCheck & "_Copy"
                nodCheck.Text = nodCheck.Key
            End If
        Else
            'If it has any children make correct them, by calling this function iteratively
            Call CorrectBranchKeys(tvwControl, nodCheck)
        End If
        Set nodCheck = nodCheck.Next
    Loop
On Error GoTo 0
Exit Sub
'Error Handler
CorrectBranchKeys_ErrorHandler:
    Exit Sub
End Sub
Public Sub CollapseBranch(ByRef tvwControl As TreeView, ByVal strKey As String)
'Created by Robin G Brown, 11/4/97
'Collapse all branches under this one
'Sub Declares
Const conSub = "CollapseBranch"
Dim nodCurrent As Node
Dim nodSibling As Node
Dim nodChild As Node
Dim intNext As Integer
    'Error Trap
    On Error GoTo CollapseBranch_ErrorHandler
    'Hide current node
    tvwControl.Nodes(strKey).Expanded = False
    'Goto first child
    Set nodCurrent = tvwControl.Nodes(strKey).Child.FirstSibling
    'If it exists hide it/goto next child
    Do While Not nodCurrent Is Nothing
        'Check for children
        Set nodChild = nodCurrent.Child
        'If it has any children make them visible, by calling this function iteratively
        If Not nodCurrent.Child Is Nothing Then
            Call CollapseBranch(tvwControl, nodCurrent.Key)
        Else
            'Make it visible
            nodCurrent.Expanded = False
        End If
        Set nodCurrent = nodCurrent.Next
    Loop
On Error GoTo 0
Exit Sub
'Error Handler
CollapseBranch_ErrorHandler:
    Exit Sub
End Sub
Public Sub ExpandBranch(ByRef tvwControl As TreeView, ByVal strKey As String)
'Created by Robin G Brown, 11/4/97
'Expand all branches under this one
'Sub Declares
Const conSub = "ExpandBranch"
Dim nodCurrent As Node
Dim nodSibling As Node
Dim nodChild As Node
Dim intNext As Integer
    'Error Trap
    On Error GoTo ExpandBranch_ErrorHandler
    'Make current node visible
    tvwControl.Nodes(strKey).EnsureVisible
    'Goto first child
    Set nodCurrent = tvwControl.Nodes(strKey).Child.FirstSibling
    'If it exists make that visible/goto next child
    Do While Not nodCurrent Is Nothing
        'Check for children
        Set nodChild = nodCurrent.Child
        'If it has any children make them visible, by calling this function recursively
        If nodCurrent.Child Is Nothing Then
            'Make it visible
            nodCurrent.EnsureVisible
        Else
            Call ExpandBranch(tvwControl, nodCurrent.Key)
        End If
        Set nodCurrent = nodCurrent.Next
    Loop
On Error GoTo 0
Exit Sub
'Error Handler
ExpandBranch_ErrorHandler:
    Exit Sub
End Sub
Public Function FindNode(ByRef nodFind As Node, ByRef tvwControl As TreeView, ByVal strNodeKey As String) As Boolean
'Created by Robin G Brown, 11/4/97
'Given a control and a key, sets nodFind to the node, sets nodFind to nothing if no match found
'Function Declares
Const conFunction = "FindNode"
Dim nodCurrent As Node
    'Error Trap
    On Error GoTo FindNode_ErrorHandler
    For Each nodCurrent In tvwControl.Nodes
        If nodCurrent.Key = strNodeKey Then
            'Set nodFind to found node thus returning a reference to the selected node
            Set nodFind = nodCurrent
            'Function was succesful so return true and exit
            FindNode = True
            Exit Function
        End If
    Next
    'If we get here without exiting then the node with strNodeKey as its key was not found _
        so set nodFind to nothing and return false
    Set nodFind = Nothing
    FindNode = False
On Error GoTo 0
Exit Function
'Error Handler
FindNode_ErrorHandler:
    'If there was an error set nodFind to nothing and return false to indicate failure
    Set nodFind = Nothing
    FindNode = False
    Exit Function
End Function
Public Function CheckForDuplicateKey(ByRef tvwControl As TreeView, ByVal strKey As String) As Integer
'Created by Robin G Brown, 11/4/97
'Checks all nodes for duplicate key, returns true if found, false otherwise
'Function Declares
Const conFunction = "CheckForDuplicateKey"
Dim nodCurrentNode As Node
    'Error Trap
    On Error GoTo CheckForDuplicateKey_ErrorHandler
    For Each nodCurrentNode In tvwControl.Nodes
        If InStr(nodCurrentNode.Key, strKey) > 0 And InStr(strKey, nodCurrentNode.Key) > 0 Then
            CheckForDuplicateKey = True
            Exit Function
        End If
    Next
    CheckForDuplicateKey = False
On Error GoTo 0
Exit Function
'Error Handler
CheckForDuplicateKey_ErrorHandler:
    CheckForDuplicateKey = False
    Exit Function
End Function
Public Sub PrintTree(ByRef tvwControl As TreeView, ByVal strKey As String)
'Created by Robin G Brown, 14/4/97
'Expand all branches under this one
'Sub Declares
Const conSub = "PrintTree"
Dim nodCurrent As Node
    'Error Trap
    On Error GoTo PrintTree_ErrorHandler
    For intCounter = 1 To tvwControl.Nodes.Count
        Debug.Print tvwControl.Nodes(intCounter).Key
    Next
On Error GoTo 0
Exit Sub
'Error Handler
PrintTree_ErrorHandler:
    Exit Sub
End Sub
Public Sub FillTreeControl(ByRef tvwControl As TreeView)
'Created by Robin G Brown, 16/4/97
'Fill the treeview control with some data for testing
'Sub Declares
Const conSub = "FillTreeControl"
Dim intTrunkCounter As Integer
Dim intBranchCounter As Integer
Dim intLeafCounter As Integer
Dim nodtemp As Node
Dim strImage As String
    On Error GoTo FillTreeControl_ErrorHandler
    'First of all clear the control
    tvwControl.Nodes.Clear
    'Check for images
    If TypeOf tvwControl.ImageList Is ImageList Then
        strImage = conDefaultTreeViewImage
    Else
        strImage = ""
    End If
    'Add a root
    Set nodtemp = tvwControl.Nodes.Add(, , "root", "Root", strImage)
    For intTrunkCounter = 1 To 2
        'Add a trunk for each tree
        Set nodtemp = tvwControl.Nodes.Add( _
            "root", _
            tvwChild, _
            "root_trunk" & intTrunkCounter, _
            "Root_Trunk" & intTrunkCounter, strImage)
        For intBranchCounter = 1 To 3
            'Add a branch
            Set nodtemp = tvwControl.Nodes.Add( _
                "root_trunk" & intTrunkCounter, _
                tvwChild, _
                "root_trunk" & intTrunkCounter & "_branch" & intBranchCounter, _
                "Root_Trunk" & intTrunkCounter & "_Branch" & intBranchCounter, strImage)
            For intLeafCounter = 1 To 4
                'Add a number of leaves
                Set nodtemp = tvwControl.Nodes.Add( _
                    "root_trunk" & intTrunkCounter & "_branch" & intBranchCounter, _
                    tvwChild, _
                    "root_trunk" & intTrunkCounter & "_branch" & intBranchCounter & "_leaf" & intLeafCounter, _
                    "Root_Trunk" & intTrunkCounter & "_Branch" & intBranchCounter & "_Leaf" & intLeafCounter, strImage)
                nodtemp.EnsureVisible
            Next
        Next
    Next
Exit Sub
FillTreeControl_ErrorHandler:
    Exit Sub
End Sub
Public Sub AddStump(ByRef tvwControl As TreeView)
'Created by Robin G Brown, 16/4/97
'Fill a treeview control with just a root node
'Sub Declares
Const conSub = "AddStump"
Dim nodtemp As Node
Dim strImage As String
    'Error Trap
    On Error GoTo AddStump_ErrorHandler
    'First of all clear the control
    tvwControl.Nodes.Clear
    If TypeOf tvwControl.ImageList Is ImageList Then
        strImage = conDefaultTreeViewImage
    Else
        strImage = ""
    End If
    'Add a root
    Set nodtemp = tvwControl.Nodes.Add(, , "root", "Root", strImage)
Exit Sub
'Error Handler
AddStump_ErrorHandler:
    Exit Sub
End Sub

