Option Strict Off

Imports System.Reflection
Imports System.IO

Public Class LateBindingSingletonWrapper

    Public Shared ReadOnly Instance As LateBindingSingletonWrapper = New LateBindingSingletonWrapper

    Private LateBoundInstance As Object

    Private Sub New()
        'Prevent anyone from instantiating this class by making New private
        Dim a As Assembly = Nothing
        Try
            a = Assembly.Load("MyPatterns.LateBinding")
        Catch ex As FileNotFoundException
            Throw
        End Try
        Dim classType As Type = a.[GetType]("MyPatterns.LateBinding.LateBound")
        LateBoundInstance = Activator.CreateInstance(classType)
    End Sub

    Public Sub SetData(value As String)
        LateBoundInstance.SetData(value)
    End Sub

    Public Function GetData() As String
        Return CType(LateBoundInstance.GetData(), String)
    End Function

End Class
