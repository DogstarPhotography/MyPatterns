Public Interface ILaunch(Of W As IWork(Of J, T, S, O), J As IJob(Of T, S, O), T As ITask(Of S, O), S As IState(Of O), O)

    Property Jobs() As Generic.List(Of J)
    Property Worker() As W

    Property ConfigFile() As String

    'Property Name() As String
    'Property Description() As String

    Sub Prepare()
    Sub Launch()
    Sub Cancel()

End Interface
