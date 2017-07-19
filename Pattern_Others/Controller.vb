Public Class Controller

    Public Event Message(ByVal Content As String)

    Public ControlKey As New Object

    Public FreeClients As New Generic.List(Of FreeClient)
    Public LockClients As New Generic.List(Of LockClient)

    Public Sub New()

        'Create some free clients
        For Counter As Integer = 1 To 3
            Dim newClient As New FreeClient(Me, "FreeClient_" & Counter.ToString)
            AddHandler newClient.WorkDone, AddressOf Client_WorkDone
            FreeClients.Add(newClient)
        Next

        'Create some lock clients
        For Counter As Integer = 1 To 3
            Dim newClient As New LockClient(Me, "LockClient_" & Counter.ToString)
            AddHandler newClient.WorkDone, AddressOf Client_WorkDone
            LockClients.Add(newClient)
        Next

    End Sub

    Public Sub Test()

        For Each curClient As FreeClient In Me.FreeClients
            curClient.DoWorkAsync()
        Next

        For Each curClient As LockClient In Me.LockClients
            curClient.DoWorkAsync()
        Next

    End Sub

    Private Sub Client_WorkDone(ByVal Name As String)

        RaiseEvent Message(Name)

    End Sub
End Class
