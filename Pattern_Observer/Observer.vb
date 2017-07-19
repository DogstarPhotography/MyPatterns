''' <summary>
''' Concrete instance of a generic observer class
''' </summary>
''' <remarks>Implements IObserver(Of T)</remarks>
Public Class Observer(Of T)
    Implements IObserver(Of T)

    Private MyObservable As IObservable(Of T)
    ''' <summary>
    ''' Register this observer with the given Observable
    ''' </summary>
    ''' <param name="TheObservable">IObservable(Of T)</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal TheObservable As IObservable(Of T))
        MyObservable = TheObservable
        MyObservable.RegisterObserver(Me)
    End Sub
    ''' <summary>
    ''' Do something with the update data
    ''' </summary>
    ''' <param name="TheData">T</param>
    ''' <remarks>This will be called by the observable object</remarks>
    Public Sub UpdateObserver(ByVal TheData As T) Implements IObserver(Of T).UpdateObserver
        'CODE_HERE
    End Sub
    ''' <summary>
    ''' Finished so remove reference to Observable
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub Finalize()
        MyObservable.RemoveObserver(Me)
        MyBase.Finalize()
    End Sub
End Class
