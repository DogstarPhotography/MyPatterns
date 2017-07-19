Public Class DataItem
    Implements IDataIdentifier, IDataUpdateStatus

    Public Field1 As String
    Public Field2 As String
    Public Field3 As String

#Region "IIdentifier"
    Private myID As Integer

    Public Property ID() As Integer Implements IDataIdentifier.ID
        Get
            Return myID
        End Get
        Set(ByVal value As Integer)
            myID = value
        End Set
    End Property
#End Region

#Region "IUpdate"
    Private myUpdateStatus As UpdateStatus

    Public Property UpdateStatus() As UpdateStatus Implements IDataUpdateStatus.UpdateStatus
        Get
            Return myUpdateStatus
        End Get
        Set(ByVal value As UpdateStatus)
            myUpdateStatus = value
        End Set
    End Property
#End Region
End Class
