Public Class Client

    Public Sub UseAdapter()

        Dim MyAdapter As ITarget = New ObjectAdapter

        MyAdapter.Request()

    End Sub
End Class
