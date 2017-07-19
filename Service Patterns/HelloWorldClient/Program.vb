Module Program

	Sub Main()
		Dim client As HelloWorldServiceClient = New HelloWorldServiceClient()
		' Use the 'client' variable to call operations on the service.
		Console.WriteLine(client.GetMessage("Robin"))

		' Always close the client.
		client.Close()
	End Sub

End Module
