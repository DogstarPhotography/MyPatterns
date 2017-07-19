Imports System.ServiceModel
Imports System.ServiceModel.Activation
Imports System.ServiceModel.Web

<ServiceContract(Namespace:="SandwichService")> _
<ServiceBehavior(IncludeExceptionDetailInFaults:=True)> _
<AspNetCompatibilityRequirements(RequirementsMode:=AspNetCompatibilityRequirementsMode.Allowed)> _
Public Class CostService

	' To use HTTP GET, add <WebGet()> attribute. (Default ResponseFormat is WebMessageFormat.Json)
	' To create an operation that returns XML,
	'     add <WebGet(ResponseFormat:=WebMessageFormat.Xml)>,
	'     and include the following line in the operation body:
	'         WebOperationContext.Current.OutgoingResponse.ContentType = "text/xml"
	<OperationContract()> _
	<WebGet(ResponseFormat:=WebMessageFormat.Json)> _
	Public Function CostOfSandwiches(quantity As Integer) As Double
		' Add your operation implementation here
		Return 1.25 * quantity
	End Function

	' Add more operations here and mark them with <OperationContract()>

End Class
