''' <summary>
''' Observable interface
''' </summary>
''' <remarks>The Observable maintains number of observers and calls their update method when NotifyObserver is called</remarks>
Public Interface IObservable
    ''' <summary>
    ''' Add an observer
    ''' </summary>
    ''' <param name="TheObserver">IObserver</param>
    ''' <remarks>How observers are stored internally is an implementation specific detail</remarks>
    Sub RegisterObserver(ByVal TheObserver As IObserver)
    ''' <summary>
    ''' Remove a previously registered observer
    ''' </summary>
    ''' <param name="TheObserver">IObserver</param>
    ''' <remarks>How observers are stored internally is an implementation specific detail</remarks>
    Sub RemoveObserver(ByVal TheObserver As IObserver)
    ''' <summary>
    ''' Call the update method on each registered observer
    ''' </summary>
    ''' <param name="TheData">String</param>
    ''' <remarks>How observers are stored internally is an implementation specific detail</remarks>
    Sub NotifyObservers(ByVal TheData As String)

End Interface
