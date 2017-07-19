Public Interface IWork(Of J As IJob(Of T, S, O), T As ITask(Of S, O), S As IState(Of O), O)

    Event JobCompleted(ByVal Worker As IWork(Of J, T, S, O), ByVal Job As J)
    Event WorkCompleted(ByVal Worker As IWork(Of J, T, S, O))

    Property Jobs() As Generic.List(Of J)

    ReadOnly Property Complete() As Boolean

    Property Name() As String
    Property Description() As String

    Sub StartWork()
    Sub StopWork()

End Interface
