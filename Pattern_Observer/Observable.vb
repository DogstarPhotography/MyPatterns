''' <summary>
''' Concrete instance of a generic observable class
''' </summary>
''' <remarks>The Observable class maintains an internal list of observers and calls their update method when NotifyObserver is called</remarks>
Public Class Observable(Of T)
    Implements IObservable(Of T)

    ''' <summary>
    ''' Internal generic list to hold observer references
    ''' </summary>
    ''' <remarks></remarks>
    Private ObserverList As List(Of IObserver(Of T))
    ''' <summary>
    ''' Standard constructor
    ''' </summary>
    ''' <remarks>Create internal list</remarks>
    Public Sub New()
        'Create the list
        ObserverList = New List(Of IObserver(Of T))
    End Sub
    ''' <summary>
    ''' Call the update method on each registered observer and pass them the supplied data
    ''' </summary>
    ''' <param name="TheData">String</param>
    ''' <remarks>The observers are stored internally in a generic list. The ObservableInstance class could instead generate the data internally</remarks>
    Public Sub NotifyObservers(ByVal TheData As T) Implements IObservable(Of T).NotifyObservers
        'Send update data to each observer in the list
        For Each curObserver As IObserver(Of T) In ObserverList
            curObserver.UpdateObserver(TheData)
        Next
    End Sub
    ''' <summary>
    ''' Add an observer
    ''' </summary>
    ''' <param name="TheObserver">IObserver(Of T)</param>
    ''' <remarks>The observers are stored internally in a generic list</remarks>
    Public Sub RegisterObserver(ByVal TheObserver As IObserver(Of T)) Implements IObservable(Of T).RegisterObserver
        ObserverList.Add(TheObserver)
    End Sub
    ''' <summary>
    ''' Remove a previously registered observer
    ''' </summary>
    ''' <param name="TheObserver">IObserver(Of T)</param>
    ''' <remarks>The observers are stored internally in a generic list</remarks>
    Public Sub RemoveObserver(ByVal TheObserver As IObserver(Of T)) Implements IObservable(Of T).RemoveObserver
        ObserverList.Remove(TheObserver)
    End Sub
End Class
