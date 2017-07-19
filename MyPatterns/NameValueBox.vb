Public Class NameValueBox
    'Event declarations
    Public Event NameValueGotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Public Event NameValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    'Variables
    Private PreviousColor As Color

    Public Property NameText() As String
        Get
            Return lblName.Text
        End Get
        Set(ByVal value As String)
            lblName.Text = value
        End Set
    End Property

    Public Property ValueText() As String
        Get
            Return txtValue.Text
        End Get
        Set(ByVal value As String)
            txtValue.Text = value
        End Set
    End Property

#Region "Events"

    Private Sub lblName_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblName.Click
        lblName.Focus()
    End Sub

    Private Sub lblName_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblName.GotFocus
        'Make focus obvious
        ShowFocus()
        'Inform controller
        RaiseEvent NameValueGotFocus(Me, New System.EventArgs)
    End Sub

    Private Sub txtValue_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtValue.GotFocus
        'Make focus obvious
        ShowFocus()
        'Inform controller
        RaiseEvent NameValueGotFocus(Me, New System.EventArgs)
    End Sub

    Private Sub txtValue_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtValue.LostFocus
        ResetFocus()
    End Sub

    Private Sub lblName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblName.LostFocus
        ResetFocus()
    End Sub

    Private Sub txtValue_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtValue.TextChanged
        RaiseEvent NameValueChanged(Me, New System.EventArgs)
    End Sub
#End Region

    Private Sub ShowFocus()
        PreviousColor = pnlTableLayout.BackColor
        pnlTableLayout.BackColor = Color.DarkGray
    End Sub

    Private Sub ResetFocus()
        pnlTableLayout.BackColor = PreviousColor
    End Sub
End Class
