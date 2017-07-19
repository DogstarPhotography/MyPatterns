Public Class _Default
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

		'In order to have a HelloWorldServiceClient class we need to add the svcutil.exe created files
		'	HelloWorldServiceRef.vb
		'	and app.config (for a web client we extract the system.servicemodel entry and add it to the web.config instead)
		'
		Dim client As HelloWorldServiceClient = New HelloWorldServiceClient()
		' Use the 'client' variable to call operations on the service.
		litHello.Text = client.GetMessage("Robin")

		' Always close the client.
		client.Close()

    End Sub

End Class