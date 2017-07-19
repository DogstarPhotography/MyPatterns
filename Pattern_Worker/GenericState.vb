Public Class GenericState(Of O As New)
    Implements IState(Of O)

    Private MyState As O
    Private MyStateID As Guid

    Public Sub New()
        MyState = New O
        MyStateID = New Guid
    End Sub

    Public Sub New(ByVal State As O, ByVal StateID As Guid)
        MyState = State
        MyStateID = StateID
    End Sub

    Public Property StateID() As System.Guid Implements IState(Of O).StateID
        Get
            Return MyStateID
        End Get
        Set(ByVal value As System.Guid)
            MyStateID = value
        End Set
    End Property

    Public Property State() As O Implements IState(Of O).State
        Get
            Return MyState
        End Get
        Set(ByVal value As O)
            MyState = value
        End Set
    End Property
End Class
