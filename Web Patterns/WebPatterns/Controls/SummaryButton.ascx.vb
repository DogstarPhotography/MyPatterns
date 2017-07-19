Public Class SummaryButton
	Inherits System.Web.UI.UserControl

	Private _Text As String
	Public Property Text() As String
		Get
			Return _Text
		End Get
		Set(ByVal value As String)
			_Text = value
			lblContent.Text = _Text.Substring(0, _Length)
		End Set
	End Property

	Private _Length As Integer = 120
	Public Property Length() As Integer
		Get
			Return _Length
		End Get
		Set(ByVal value As Integer)
			_Length = value
			lblContent.Text = _Text.Substring(0, _Length)
		End Set
	End Property


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	End Sub

End Class