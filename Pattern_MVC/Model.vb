''' <summary>
''' The model class is Observable by views or controllers
''' </summary>
''' <remarks></remarks>
Public Class Model
    Implements IObservable
    Implements IModel

#Region "IObservable implementation"
    ''' <summary>
    ''' Internal generic list too hold observer references
    ''' </summary>
    ''' <remarks></remarks>
    Private ObserverList As List(Of IObserver)
    ''' <summary>
    ''' Standard constructor
    ''' </summary>
    ''' <remarks>Create internal list</remarks>
    Public Sub New()
        'Create the list
        ObserverList = New List(Of IObserver)
    End Sub
    ''' <summary>
    ''' Call the update method on each registered observer and pass them the supplied data
    ''' </summary>
    ''' <param name="TheData">String</param>
    ''' <remarks>The observers are stored internally in a generic list. The subjectinstance class could instead generate the data internally</remarks>
    Public Sub NotifyObservers(ByVal TheData As String) Implements IObservable.NotifyObservers
        'Send update data to each observer in the list
        For Each curObserver As IObserver In ObserverList
            curObserver.UpdateObserver(TheData)
        Next
    End Sub
    ''' <summary>
    ''' Add an observer
    ''' </summary>
    ''' <param name="TheObserver">IObserver</param>
    ''' <remarks>The observers are stored internally in a generic list</remarks>
    Public Sub RegisterObserver(ByVal TheObserver As IObserver) Implements IObservable.RegisterObserver
        ObserverList.Add(TheObserver)
    End Sub
    ''' <summary>
    ''' Remove a previously registered observer
    ''' </summary>
    ''' <param name="TheObserver">IObserver</param>
    ''' <remarks>The observers are stored internally in a generic list</remarks>
    Public Sub RemoveObserver(ByVal TheObserver As IObserver) Implements IObservable.RemoveObserver
        ObserverList.Remove(TheObserver)
    End Sub
#End Region

#Region "Internal model implementation"
    'CODE_HERE_TO_DO
#End Region

    Public Sub UpdateState(ByVal Dat As String) Implements IModel.UpdateState

    End Sub
End Class
