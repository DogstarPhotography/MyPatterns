Public Class HelloWorldService
	Implements IHelloWorldService

	Public Function GetMessage(name As String) As String Implements IHelloWorldService.GetMessage
		Return "Hello " & name
	End Function
End Class
