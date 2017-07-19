<ComClass(SampleAPI.ClassId, SampleAPI.InterfaceId, SampleAPI.EventsId)> _
Public Class SampleAPI

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "3ed8cbf2-a118-49ea-a8cf-46b4e2e8dfc8"
    Public Const InterfaceId As String = "4ec1e83b-5ff8-4789-a755-8f9b652baf08"
    Public Const EventsId As String = "d6ca70bb-fe77-4559-9ac8-d3470062bb58"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Private myValue As String = ""

    Public Sub MethodA()
        myValue = "TEST"
    End Sub

    Public Sub MethodB(parameter As String)
        myValue = parameter
    End Sub

    Public Function MethodC() As String
        Return myValue
    End Function

End Class


