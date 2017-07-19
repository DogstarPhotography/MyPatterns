Imports System.IO

Public Interface IJob(Of T As ITask(Of S, O), S As IState(Of O), O)

    Enum JobStatus
        Waiting
        InProgress
        Complete
        Problem
    End Enum

    Event Updated(ByVal TheJob As IJob(Of T, S, O))
    Event Completed(ByVal TheJob As IJob(Of T, S, O))
    'Event Problem()

    Property Tasks() As Collections.Generic.Queue(Of T)

    ReadOnly Property Status() As JobStatus

    'Property ConfigFile() As String

    'Property JobID() As Guid
    Property Name() As String
    Property Description() As String

    Sub StartJob()
    Sub StopJob()

End Interface
