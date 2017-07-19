''' <summary>
''' Observer interface
''' </summary>
''' <remarks></remarks>
Public Interface IObserver
    ''' <summary>
    ''' Use the update data supplied
    ''' </summary>
    ''' <param name="TheData">String</param>
    ''' <remarks>What observers do with the data is an implementation specific detail</remarks>
    Sub UpdateObserver(ByVal TheData As String)

End Interface
