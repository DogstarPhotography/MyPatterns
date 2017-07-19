Imports System.ComponentModel

''' <summary>
''' Simple generic worker that processes a list of jobs in no particular order
''' </summary>
''' <typeparam name="J">IJob(Of T, S, O)</typeparam>
''' <typeparam name="T">ITask(Of S, O)</typeparam>
''' <typeparam name="S">IState(Of O)</typeparam>
''' <typeparam name="O">Object</typeparam>
''' <remarks></remarks>
Public Class GenericWorker(Of J As IJob(Of T, S, O), T As ITask(Of S, O), S As IState(Of O), O)
    Implements IWork(Of J, T, S, O)

    Public Event JobCompleted(ByVal Worker As IWork(Of J, T, S, O), ByVal Job As J) Implements IWork(Of J, T, S, O).JobCompleted
    Public Event WorkCompleted(ByVal Worker As IWork(Of J, T, S, O)) Implements IWork(Of J, T, S, O).WorkCompleted

    Private MyJobs As Generic.List(Of J)
    Private CompleteFlag As Boolean = False

    Public Sub StartWork() Implements IWork(Of J, T, S, O).StartWork
        DoWork()
    End Sub

    Public Sub StopWork() Implements IWork(Of J, T, S, O).StopWork
        'Stop each job
        For Each curJob As IJob(Of T, S, O) In MyJobs
            curJob.StopJob()
        Next
    End Sub

    Private Sub DoWork()
        'Do each of the jobs in the list in no particular order
        For Each curJob As IJob(Of T, S, O) In MyJobs
            AddHandler curJob.Completed, AddressOf JobCompletedHandler
            curJob.StartJob()
        Next
    End Sub

    Private Sub JobCompletedHandler(ByVal TheJob As IJob(Of T, S, O))
        'Inform the consumer
        RaiseEvent JobCompleted(Me, CType(TheJob, J))
        'Check each job
        For Each curJob As IJob(Of T, S, O) In MyJobs
            If curJob.Status <> IJob(Of T, S, O).JobStatus.Complete Then Exit Sub
        Next
        'If all jobs are complete set the flag and raise the event
        CompleteFlag = True
        RaiseEvent WorkCompleted(Me)
    End Sub

#Region "Properties"
    Public Property Jobs() As System.Collections.Generic.List(Of J) Implements IWork(Of J, T, S, O).Jobs
        Get
            Return MyJobs
        End Get
        Set(ByVal value As System.Collections.Generic.List(Of J))
            MyJobs = value
        End Set
    End Property

    Public ReadOnly Property Complete() As Boolean Implements IWork(Of J, T, S, O).Complete
        Get
            Return CompleteFlag
        End Get
    End Property

    Public Property Description() As String Implements IWork(Of J, T, S, O).Description
        Get
            Return Me.ToString
        End Get
        Set(ByVal value As String)

        End Set
    End Property

    Public Property Name() As String Implements IWork(Of J, T, S, O).Name
        Get
            Return Me.ToString
        End Get
        Set(ByVal value As String)

        End Set
    End Property
#End Region
End Class
