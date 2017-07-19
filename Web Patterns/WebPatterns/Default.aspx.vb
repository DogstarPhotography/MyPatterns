Public Class _Default
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

		HttpContext.Current.Response.Write("""Blah""")

    End Sub

End Class