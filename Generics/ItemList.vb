''' <summary>
''' Class that inherits from a simple list. Demonstrates how to implement a constructor that takes a list.
''' </summary>
''' <remarks></remarks>
Public Class ItemList
    Inherits List(Of String)
    ''' <summary>
    ''' Default constructor
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

    End Sub
    ''' <summary>
    ''' We want a constructor that takes a list of strings.
    ''' </summary>
    ''' <param name="items"></param>
    ''' <remarks></remarks>
    Public Sub New(items As List(Of String))
        For Each Item As String In items
            Me.Add(Item)
        Next
    End Sub

End Class
