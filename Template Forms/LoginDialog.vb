Public Class LoginDialog

    Public Property Identity As String = ""
    Public Property Password As String = ""

    Public Shadows Function ShowDialog() As DialogResult
        txtIdentity.Text = Me.Identity
        txtPassword.Text = Me.Password
        Return MyBase.ShowDialog
    End Function

    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        Me.Identity = txtIdentity.Text
        Me.Password = txtPassword.Text
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Identity = txtIdentity.Text
        Me.Password = ""
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

End Class