Imports System.IO

Public Interface IJob(Of T As ITask(Of S, O), S As IState(Of O), O)

    'Event Progress()
    Event Completed(ByVal TheJob As IJob(Of T, S, O))
    'Event Problem()

    Property Tasks() As Collections.Generic.Queue(Of T)

    ReadOnly Property Complete() As Boolean

    'Property ConfigFile() As String

    'Property JobID() As Guid
    'Property Name() As String
    'Property Description() As String

    Sub StartJob()
    'Sub Cancel()

End Interface
