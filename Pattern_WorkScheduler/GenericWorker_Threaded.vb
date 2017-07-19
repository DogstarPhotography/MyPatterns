Imports System.ComponentModel

''' <summary>
''' Simple generic worker that processes a list of jobs in no particular order
''' </summary>
''' <typeparam name="J">IJob(Of T, S, O)</typeparam>
''' <typeparam name="T">ITask(Of S, O)</typeparam>
''' <typeparam name="S">IState(Of O)</typeparam>
''' <typeparam name="O">Object</typeparam>
''' <remarks>Multithreaded version of GenericWorker for improved responsiveness</remarks>
Public Class GenericWorker_Threaded(Of J As IJob(Of T, S, O), T As ITask(Of S, O), S As IState(Of O), O)
    Implements IWork(Of J, T, S, O)

    Public Event JobCompleted(ByVal Worker As IWork(Of J, T, S, O), ByVal Job As J) Implements IWork(Of J, T, S, O).JobCompleted
    Public Event WorkCompleted(ByVal Worker As IWork(Of J, T, S, O)) Implements IWork(Of J, T, S, O).WorkCompleted

    Private MyJobs As Generic.List(Of J)
    Private CompleteFlag As Boolean = False

    Public Sub StartWork() Implements IWork(Of J, T, S, O).StartWork
        'Start the asynchronous operation.
        MyBackgroundWorker = New BackgroundWorker
        MyBackgroundWorker.RunWorkerAsync()
    End Sub

    Public Sub StopWork() Implements IWork(Of J, T, S, O).StopWork
        'Stop the backgroundworker
        MyBackgroundWorker.CancelAsync()
    End Sub

#Region "BackgroundWorker"
    'Background worker will do the actual work on a different thread
    Private WithEvents MyBackgroundWorker As BackgroundWorker
    Private numCompletedJobs As Integer

    Private Sub MyBackgroundWorker_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles MyBackgroundWorker.DoWork
        'Count number of completed jobs for reporting progress
        numCompletedJobs = 0
        'Do each of the jobs in the list in no particular order
        For Each curJob As IJob(Of T, S, O) In MyJobs
            AddHandler curJob.Completed, AddressOf JobCompletedHandler
            curJob.StartJob() 'Blocking or async call
        Next
        'Once the work is done MyBackgroundWorker_RunWorkerCompleted will be raised so there is no need to do anything here
    End Sub

    Private Sub JobCompletedHandler(ByVal TheJob As IJob(Of T, S, O))
        'Report progress
        numCompletedJobs = numCompletedJobs + 1
        MyBackgroundWorker.ReportProgress(CInt((numCompletedJobs / MyJobs.Count) * 100), TheJob)
    End Sub

    Private Sub MyBackgroundWorker_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles MyBackgroundWorker.ProgressChanged
        'Report progress to consumer
        RaiseEvent JobCompleted(Me, CType(e.UserState, J))
    End Sub

    Private Sub MyBackgroundWorker_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles MyBackgroundWorker.RunWorkerCompleted
        'As all jobs are complete set the flag and raise the event
        CompleteFlag = True
        RaiseEvent WorkCompleted(Me)
    End Sub
#End Region

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
