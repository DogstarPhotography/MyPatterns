Public Class MyDictionaryEditor

    'MyDictionary object
    Private TheDictionary As MyDictionary
    'Variable to store changed state of mydictionary object
    Private FileChanged As Boolean = False
    'Variable to store key of currently selected item
    Private CurrentItemKey As String = ""

    Private Sub MyDictionaryEditor_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Clear design time controls
        pnlItems.Controls.Clear()
    End Sub

#Region "File Menu"
    'Code to implement file menu actions
    Private Sub mnuFileNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileNew.Click
        'Check for save and cancel
        If CancelForSave() = True Then Exit Sub
        'Create a new mydictionary object
        TheDictionary = New MyDictionary
        'Add an item
        'mnuEditAdd_Click(Nothing, Nothing)
        'Display
        DisplayMyDictionary()
    End Sub

    Private Sub mnuFileOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileOpen.Click
        'Check for save and cancel
        If CancelForSave() = True Then Exit Sub
        'Open a  file
        If dlgFileOpen.ShowDialog() = Windows.Forms.DialogResult.OK Then
            'Create new object
            TheDictionary = New MyDictionary
            'We really are going to save the file
            TheDictionary.LoadXML(dlgFileOpen.FileName)
            'Set the changed flag to false
            FileChanged = False
            'Display
            DisplayMyDictionary()
        End If
    End Sub

    Private Sub mnuFileSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileSave.Click
        'Save the mydictionary object
        If dlgFileSave.ShowDialog() = Windows.Forms.DialogResult.OK Then
            'We really are going to save the file
            TheDictionary.SaveXML(dlgFileSave.FileName)
            'Set the changed flag to false
            FileChanged = False
            'Display
            DisplayMyDictionary()
        End If
    End Sub

    Private Sub mnuFileExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileExit.Click
        'Check for save or cancel
        If FileChanged = True Then
            If MessageBox.Show("File not saved, exit anyway?", "Exit", MessageBoxButtons.OKCancel, MessageBoxIcon.None, MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.Cancel Then Exit Sub
        End If
        'Close
        Me.Close()
    End Sub

    Private Function CancelForSave() As Boolean
        'Exit if file has not changed
        If FileChanged = False Then
            Return False
        Else
            'See if the user wants to save the changes
            If MessageBox.Show("File is not saved, save now?", "Save?", MessageBoxButtons.OKCancel, MessageBoxIcon.None, MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.Cancel Then
                'Cancel save
                Return False
            Else
                If dlgFileSave.ShowDialog() = Windows.Forms.DialogResult.Cancel Then
                    Return False
                Else
                    'We really are going to save the file
                    TheDictionary.SaveXML(dlgFileSave.FileName)
                    'Set the changed flag to false
                    FileChanged = False
                    Return True
                End If
            End If
        End If
    End Function
#End Region

#Region "Edit Menu"
    'Code to implement edit menu actions
    Private Sub mnuEditAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEditAdd.Click
        Static ctrNew As Integer = 0
        If TheDictionary Is Nothing Then Exit Sub
        'Add a new item as the current item
        ctrNew = ctrNew + 1
        CurrentItemKey = "New" & ctrNew.ToString
        TheDictionary.Items.Add(CurrentItemKey, "")
        'Display
        FileChanged = True
        DisplayMyDictionary()
    End Sub

    Private Sub mnuEditDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEditDelete.Click
        If TheDictionary Is Nothing Then Exit Sub
        'Check with user
        If MessageBox.Show("Are you sure?", "Delete?", MessageBoxButtons.OKCancel, MessageBoxIcon.None, MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.OK Then
            'Delete current item
            TheDictionary.Items.Remove(CurrentItemKey)
            'Clear current item key
            CurrentItemKey = ""
            'Display
            FileChanged = True
            DisplayMyDictionary()
        End If
    End Sub

    Private Sub mnuEditChange_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEditChange.Click
        If TheDictionary Is Nothing Then Exit Sub
        'Change the name/key of the item
        If dlgTextEdit.ShowEdit(CurrentItemKey, "Edit Name", "Enter the new name") = Windows.Forms.DialogResult.OK Then
            TheDictionary.Items.Add(dlgTextEdit.EditedText, TheDictionary.Items(CurrentItemKey))
            TheDictionary.Items.Remove(CurrentItemKey)
            CurrentItemKey = dlgTextEdit.EditedText
            dlgTextEdit.Dispose()
            DisplayMyDictionary()
        End If
    End Sub
#End Region

    Private Sub DisplayMyDictionary()
        'Clear controls!
        pnlItems.Controls.Clear()
        'Display the contents of the mydictionary object
        For Each curItem As KeyValuePair(Of String, String) In TheDictionary.Items
            'Create a namevaluebox control 
            Dim nvbNew As New NameValueBox
            'Set properties
            nvbNew.NameText = curItem.Key
            nvbNew.ValueText = curItem.Value
            'nvbNew.Anchor = CType(AnchorStyles.Left + AnchorStyles.Right, AnchorStyles)
            'Add event handlers!
            AddHandler nvbNew.NameValueGotFocus, AddressOf nvbItem_GotFocus
            AddHandler nvbNew.NameValueChanged, AddressOf nvbItem_Changed
            'Add it to the table layout panel
            pnlItems.Controls.Add(nvbNew)
            'pnlItems.RowCount = pnlItems.Controls.Count
        Next
        'Set focus
        For Each curControl As NameValueBox In pnlItems.Controls
            If curControl.NameText = CurrentItemKey Then
                curControl.Focus()
            End If
        Next
    End Sub

#Region "NameValueBox"
    'Code to handle edits on namevaluebox control
    Private Sub nvbItem_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        'We will be using this procedure to handle several controls, 
        ' so make sure we refer only to the sender
        Dim nvbSender As NameValueBox = CType(sender, NameValueBox)
        'Set the current item key
        CurrentItemKey = nvbSender.NameText
    End Sub

    Private Sub nvbItem_Changed(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Either the name or value of nvbItem changed
        'We will be using this procedure to handle several controls, 
        ' so make sure we refer only to the sender
        Dim nvbSender As NameValueBox = CType(sender, NameValueBox)
        'Update value in TheDictionary
        TheDictionary.Items(CurrentItemKey) = nvbSender.ValueText
    End Sub
#End Region
End Class