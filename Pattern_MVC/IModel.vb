''' <summary>
''' IModel inherits an Observable interface
''' </summary>
''' <remarks>The model maintains number of observers and calls their update method when NotifyObserver is called</remarks>
Public Interface IModel
    Inherits IObservable
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Dat"></param>
    ''' <remarks></remarks>
    Sub UpdateState(ByVal Dat As String)

End Interface
