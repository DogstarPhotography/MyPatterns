Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols

Namespace HelloWorldAjaxWebServiceSite.WebService
	' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
	' <System.Web.Script.Services.ScriptService()> _
	<WebService(Namespace:="http://tempuri.org/")> _
	<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
	<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
	Public Class HelloWorld
		Inherits System.Web.Services.WebService

		'http://localhost/HelloWorldService/HelloWorld.asmx/HelloWorld
		<WebMethod()> _
		Public Function HelloWorld() As String
			Return "Hello World"
		End Function

		'Uncomment the following line if using designed components 
		'InitializeComponent(); 

		'http://localhost/HelloWorldService/HelloWorld.asmx/Greetings
		<WebMethod()> _
		Public Function Greetings() As String
			Dim serverTime As String = _
			 String.Format("Current date and time: {0}.", DateTime.Now)
			Dim greet As String = "Hello World. <br/>" + serverTime
			Return greet
		End Function 'Greetings
	End Class
End Namespace
