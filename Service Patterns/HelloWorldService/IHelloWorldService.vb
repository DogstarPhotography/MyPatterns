<ServiceContract()>
Public Interface IHelloWorldService

	<OperationContract()>
	Function GetMessage(name As String) As String

End Interface
