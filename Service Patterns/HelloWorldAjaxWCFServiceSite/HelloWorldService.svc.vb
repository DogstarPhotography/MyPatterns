'The service code contained in the HelloWorldService.svc.vb file.
Imports System.ServiceModel
Imports System.ServiceModel.Activation
Imports System.ServiceModel.Web

'DO NOT add a Namespace here  - it will mess up the namespace entry in the ServiceContract attribute
'Namespace HelloWorldAjaxWCFServiceWebApp

'Note that the precise attributes on the class and the functions are required in order to get this working
<ServiceContract(Namespace:="HelloWorldAjaxWCFServiceWebApp")>
<AspNetCompatibilityRequirements(RequirementsMode:=AspNetCompatibilityRequirementsMode.Allowed)>
Public Class HelloWorldService

	' To use HTTP GET, add <WebGet()> attribute. (Default ResponseFormat is WebMessageFormat.Json)
	' To create an operation that returns XML,
	'     add <WebGet(ResponseFormat:=WebMessageFormat.Xml)>,
	'     and include the following line in the operation body:
	'         WebOperationContext.Current.OutgoingResponse.ContentType = "text/xml"

	<OperationContract()>
	Public Function Echo(ByVal data As String) As String
		Return data
    End Function

	'UriTemplate:="auth", BodyStyle:=WebMessageBodyStyle.Bare, RequestFormat:=WebMessageFormat.Xml, ResponseFormat:=WebMessageFormat.Xml
    <OperationContract()>
    <WebInvoke(Method:="POST")>
    Public Function HelloWorld(ByVal Name As String) As String
        'At this point we can do pretty much anything we like with the data 
        '   but saving it to a database is the most likely option.
        If Name.Length > 100 Then
            Return "Hello " & Name.Substring(0, 100)
        Else
            Return "Hello " & Name
        End If
    End Function

	' Add more operations here and mark them with <OperationContract()>

End Class
'End Namespace
