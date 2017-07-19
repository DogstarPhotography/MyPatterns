Public Class frmBoilerplate

#Region "File Menu and Dialogs"
    'Basic file open and save dialogs

    Private Sub mnuFileLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileLoad.Click

        'Set a default filename
        dlgFileOpen.DefaultExt = "txt"
        dlgFileOpen.FileName = "MyFile"

        'Get a file name
        If dlgFileOpen.ShowDialog() = Windows.Forms.DialogResult.OK Then

            'Display file in form caption
            Me.Text = Me.Text & " - " & dlgFileOpen.FileName

            'Open the file
            'CODE_HERE

        End If

    End Sub

    Private Sub mnuFileSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileSave.Click

        'Set a default filename
        dlgFileSave.DefaultExt = "txt"
        dlgFileSave.FileName = "MyFile"

        'Get a file name
        If dlgFileSave.ShowDialog() = Windows.Forms.DialogResult.OK Then

            'Save the file
            'CODE_HERE

            'Display file in form caption
            Me.Text = Me.Text & " - " & dlgFileSave.FileName

        End If

    End Sub

    Private Sub mnuFileExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileExit.Click

        'Close the form
        Me.Close()

    End Sub
#End Region

#Region "Thread Safe Control Update"
    ' This method demonstrates a pattern for making thread-safe
    ' calls on a Windows Forms control. 
    '
    ' If the calling thread is different from the thread that
    ' created the TextBox control, this method creates a
    ' SetTextCallback and calls itself asynchronously using the
    ' Invoke method.
    ' or if the calling thread is the same as the thread that created
    ' the TextBox control, the Text property is set directly. 
    '
    'INSTRUCTIONS: Find and replace 'ctlControl' with name of control
    '
    ' This delegate enables asynchronous calls for setting
    ' the text property on a TextBox control.
    Delegate Sub InvokectlControlCallback(ByVal NewValue As String)

    Private Sub InvokectlControl(ByVal NewValue As String)
        ' InvokeRequired required compares the thread ID of the
        ' calling thread to the thread ID of the creating thread.
        ' If these threads are different, it returns true.
        If Me.ctlControl.InvokeRequired Then
            Dim MyCallback As New InvokectlControlCallback(AddressOf InvokectlControl)
            Invoke(MyCallback, New Object() {NewValue})
        Else
            ctlControl.Text = NewValue
        End If
    End Sub
#End Region

End Class