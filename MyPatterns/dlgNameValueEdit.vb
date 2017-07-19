Public Class dlgNameValueEdit

    Public Function ShowEdit(Optional ByVal Name As String = "", Optional ByVal Value As String = "", Optional ByVal Title As String = "", Optional ByVal Instructions As String = "", Optional ByVal SaveText As String = "&Save", Optional ByVal CancelText As String = "&Cancel") As DialogResult

        'Grab inputs
        txtEditName.Text = Name
        txtEditValue.Text = Value
        Me.Text = Title
        lblInstructions.Text = Instructions
        btnSave.Text = SaveText
        btnCancel.Text = CancelText

        'Show modally
        Return Me.ShowDialog()

    End Function

    Public ReadOnly Property EditedName() As String
        Get
            Return txtEditName.Text
        End Get
    End Property

    Public ReadOnly Property EditedValue() As String
        Get
            Return txtEditValue.Text
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