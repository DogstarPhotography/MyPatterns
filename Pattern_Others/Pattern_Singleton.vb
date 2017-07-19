''' <summary>
''' 
''' </summary>
''' <remarks>NotInheritable prevents this class from being used</remarks>
Public NotInheritable Class Pattern_Singleton
    ''' <summary>
    ''' Shared internal variables
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared TheSingleton As Pattern_Singleton
    Private Shared SyncObject As New Object
    ''' <summary>
    ''' Making the Constructor private prevents instances of this class being created
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub New()
        'No code required here
    End Sub
    ''' <summary>
    ''' Shared function returns a single instance of this class
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Instance() As Pattern_Singleton
        'Need to prevent multiple threads from accessing this at once
        SyncLock SyncObject
            'Create a new singleton if we don't have one
            If TheSingleton Is Nothing Then
                TheSingleton = New Pattern_Singleton
            End If
        End SyncLock
        'Return the only instance
        Return TheSingleton
    End Function

End Class
