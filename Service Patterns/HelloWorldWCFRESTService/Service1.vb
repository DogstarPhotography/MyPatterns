Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.ServiceModel
Imports System.ServiceModel.Activation
Imports System.ServiceModel.Web
Imports System.Text

Namespace HelloWorldWCFRESTService
    ' Start the service and browse to http://<machine_name>:<port>/Service1/help to view the service's generated help page
    ' NOTE: By default, a new instance of the service is created for each call; change the InstanceContextMode to Single if you want
    ' a single instance of the service to process all calls.	
    ' NOTE: If the service has been renamed, remember to update the global.asax.vb file
    <ServiceContract()> 
    <AspNetCompatibilityRequirements(RequirementsMode := AspNetCompatibilityRequirementsMode.Allowed)>
    <ServiceBehavior(InstanceContextMode := InstanceContextMode.PerCall)>
    Public Class Service1
        ' TODO: Implement the collection resource that will contain the SampleItem instances
        
        <WebGet(UriTemplate := "")>
        Public Function GetCollection() As List(Of SampleItem)
            ' TODO: Replace the current implementation to return a collection of SampleItem instances
            Return New List(Of SampleItem) From {New SampleItem With {.Id = 1, .StringValue = "Hello"}}
        End Function
        
        <WebInvoke(UriTemplate := "", Method := "POST")>
        Public Function Create(ByVal instance As SampleItem) As SampleItem
            ' TODO: Add the new instance of SampleItem to the collection
            Throw New NotImplementedException()
        End Function
        
        <WebGet(UriTemplate := "{id}")>
        Public Function [Get](ByVal id As String) As SampleItem
            ' TODO: Return the instance of SampleItem with the given id
            Throw New NotImplementedException()
        End Function
        
        <WebInvoke(UriTemplate := "{id}", Method := "PUT")>
        Public Function Update(ByVal id As String, ByVal instance As SampleItem) As SampleItem
            ' TODO: Update the given instance of SampleItem in the collection
            Throw New NotImplementedException()
        End Function
        
        <WebInvoke(UriTemplate := "{id}", Method := "DELETE")>
        Public Sub Delete(ByVal id As String)
            ' TODO: Remove the instance of SampleItem with the given id from the collection
            Throw New NotImplementedException()
        End Sub
        
    End Class
End Namespace