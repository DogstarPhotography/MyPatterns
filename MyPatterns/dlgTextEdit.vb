Public Class dlgTextEdit

    Public Function ShowEdit(Optional ByVal Value As String = "", Optional ByVal Title As String = "", Optional ByVal Instructions As String = "", Optional ByVal SaveText As String = "&Save", Optional ByVal CancelText As String = "&Cancel") As DialogResult

        'Grab inputs
        txtEdit.Text = Value
        Me.Text = Title
        lblInstructions.Text = Instructions
        btnSave.Text = SaveText
        btnCancel.Text = CancelText

        'Show modally
        Return Me.ShowDialog()

    End Function

    Public ReadOnly Property EditedText() As String
        Get
            Return txtEdit.Text
        End Get
    End Property

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Me.DialogResult = Windows.Forms.DialogResult.OK
        Me.Hide()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Hide()
    End Sub
End Class