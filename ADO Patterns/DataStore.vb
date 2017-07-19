Imports System.Data

''' <summary>
''' Singleton store for data components
''' </summary>
''' <remarks></remarks>
Public NotInheritable Class DataStore(Of Connection As {System.Data.Common.DbConnection, New}, Command As {System.Data.Common.DbCommand, New}, DataReader As {System.Data.Common.DbDataReader})

#Region "Data Components"

    Public ItemData As DataComponent(Of Connection, Command, DataReader)

#End Region

#Region "Singleton"
    Private Shared myInstance As DataStore(Of Connection, Command, DataReader)
    Private Shared mySyncObject As New Object
    Private Shared myConnectionString As String '= "data source=C:\Rameses\Services\gpstracking\YYYY-MM-DD\GPS-YYYY-MM-DD.s3db"

    Private Sub New()
        'Prevent anyone from instantiating this class by making New private
    End Sub

    Public Shared ReadOnly Property Instance() As DataStore(Of Connection, Command, DataReader)
        Get
            SyncLock mySyncObject
                If myInstance Is Nothing Then
                    myInstance = New DataStore(Of Connection, Command, DataReader)
                    'Create data components now we have an instance
                    myInstance.ItemData = New DataComponent(Of Connection, Command, DataReader)(myInstance)
                End If
            End SyncLock
            Return myInstance
        End Get
    End Property

    Public Property ConnectionString() As String
        Get
            Return myConnectionString
        End Get
        Set(ByVal value As String)
            myConnectionString = value
        End Set
    End Property
#End Region

End Class
