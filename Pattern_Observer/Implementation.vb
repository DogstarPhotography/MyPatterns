Friend Class Implementation

    Public Sub Test()

        Dim MyObserver As IObserver(Of String)
        Dim MySubject As IObservable(Of String)

        'Create a subject
        MySubject = New Observable(Of String)

        'Create an observer and link it to the subject
        MyObserver = New Observer(Of String)(MySubject)

        'Tell the subject to update the observers
        MySubject.NotifyObservers("SOME_DATA_HERE")

    End Sub

End Class
