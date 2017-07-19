''' <summary>
''' Generic Observer interface
''' </summary>
''' <remarks></remarks>
Public Interface IObserver(Of T)
    ''' <summary>
    ''' Use the update data supplied
    ''' </summary>
    ''' <param name="TheData">T</param>
    ''' <remarks>This will be called by the observable class, what observers do with the data is an implementation specific detail</remarks>
    Sub UpdateObserver(ByVal TheData As T)

End Interface
''' <summary>
''' Simple string specific Observer interface
''' </summary>
''' <remarks></remarks>
Public Interface IObserver
    ''' <summary>
    ''' Use the update data supplied
    ''' </summary>
    ''' <param name="TheData">String</param>
    ''' <remarks>This will be called by the observable class, what observers do with the data is an implementation specific detail</remarks>
    Sub UpdateObserver(ByVal TheData As String)

End Interface