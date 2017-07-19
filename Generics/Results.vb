''' <summary>
''' Can be used to return a result from a procedure where several bits of data need to be tied together, 
''' such as wether the procedure failed or succeeded, what the result code was, any message, and any exception details.
''' </summary>
''' <remarks></remarks>
Public Class ProcedureResult

    Public Result As Boolean
    Public Code As Integer
    Public Message As String
    Public Exception As Exception

    Public Sub New(ByVal Result As Boolean, ByVal Code As Integer, ByVal Message As String, ByVal Exception As Exception)
        Me.Result = Result
        Me.Code = Code
        Me.Message = Message
        Me.Exception = Exception
    End Sub
End Class
''' <summary>
''' ProcedureResult that includes an object to allow for passing of untyped data out of a procedure
''' </summary>
''' <remarks></remarks>
Public Class FunctionResult
    Inherits ProcedureResult

    Public Data As Object

    Public Sub New(ByVal Result As Boolean, ByVal Code As Integer, ByVal Message As String, ByVal Exception As Exception, ByVal Data As Object)
        MyBase.New(Result, Code, Message, Exception)
        Me.Data = Data
    End Sub
End Class
''' <summary>
''' ProcedureResult that includes a generic to allow for passing of typed data out of a procedure
''' </summary>
''' <typeparam name="T"></typeparam>
''' <remarks></remarks>
Public Class FunctionResult(Of T)
    Inherits ProcedureResult

    Public Data As T

    Public Sub New(ByVal Result As Boolean, ByVal Code As Integer, ByVal Message As String, ByVal Exception As Exception, ByVal Data As T)
        MyBase.New(Result, Code, Message, Exception)
        Me.Data = Data
    End Sub
End Class
