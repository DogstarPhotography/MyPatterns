Namespace Models

    Public Class Task

        Property Name As String = "Unknown"
        Property Description As String = "No description available."
        Property Selected As Boolean = True
        Property Status As String = "Not started."
        Property Progress As Integer = 0

    End Class

    Public Class TaskRequestViewModel

        Property Tasks As New List(Of Task)

    End Class

    Public Class TaskProgressViewModel

        Property Tasks As New List(Of Task)
        Property Finished As Boolean = False

        Public Sub New()

        End Sub

        Public Sub New(source As TaskRequestViewModel)
            For Each item As Task In source.Tasks
                Me.Tasks.Add(item)
            Next
        End Sub

    End Class

End Namespace
