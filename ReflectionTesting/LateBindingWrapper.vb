Option Strict Off 'Required for late binding

Imports System.Reflection
Imports System.IO

Public Class LateBindingWrapper
    ''' <summary>
    ''' Call this first to get an instance of the class we want to use
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetInstance() As Object
        Dim a As Assembly = Nothing
        Try
            a = Assembly.Load("MyPatterns.LateBinding")
        Catch ex As FileNotFoundException
            Throw
        End Try
        Dim classType As Type = a.[GetType]("MyPatterns.LateBinding.LateBound")
        Dim instance As Object = Activator.CreateInstance(classType)
        Return instance
    End Function
    ''' <summary>
    ''' Call this to execute METHOD.
    ''' </summary>
    ''' <param name="instance"></param>
    ''' <remarks></remarks>
    Public Shared Sub SetData(ByRef instance As Object, value As String)
        instance.SetData(value)
    End Sub
    ''' <summary>
    ''' Call this to execute METHOD.
    ''' </summary>
    ''' <param name="instance"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetData(ByRef instance As Object) As String
        Return instance.GetData()
    End Function

#Region "Boilerplate"
    ' ''' <summary>
    ' ''' Call this to execute METHOD.
    ' ''' </summary>
    ' ''' <param name="instance"></param>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Public Shared Function METHOD(ByRef instance As Object, parameter As String) As String
    '    Return instance.METHOD(parameter)
    'End Function

#End Region

End Class
