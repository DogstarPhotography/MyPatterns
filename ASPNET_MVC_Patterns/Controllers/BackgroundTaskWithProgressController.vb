Imports System.Web.Mvc
Imports ASPNET_MVC_Patterns.Models
Imports ASPNET_MVC_Patterns.Infrastructure
Imports System.Threading

Namespace Controllers

    Public Class BackgroundTaskWithProgressController
        Inherits Controller

        Private ProgressID As String = "TaskProgressViewModel"
        ''' <summary>
        ''' The start page simply fetches the data and displays it.
        ''' </summary>
        Function Start() As ActionResult
            'Create data for demonstration purposes
            Dim data As New TaskRequestViewModel
            With data.Tasks
                .Add(New Task With {.Name = "Task 1", .Description = "This is task 1"})
                .Add(New Task With {.Name = "Task 2", .Description = "This is task 2"})
                .Add(New Task With {.Name = "Task 3", .Description = "This is task 3"})
            End With
            Return View(data)
        End Function
        ''' <summary>
        ''' When the user clicks on the submit button in the start page the TaskInitiation function is called which invokes a task in the background and redirects to the TaskMonitor page.
        ''' </summary>
        <HttpPost>
        Function TaskInitiation(data As TaskRequestViewModel) As ActionResult
            Dim task As New TaskDelegate(AddressOf TaskWorker)
            task.BeginInvoke(HttpContext.Session, data, Nothing, Nothing)
            Return RedirectToAction("TaskMonitor")
        End Function
        ''' <summary>
        ''' Delegate for the background worker; used to call it on a different thread.
        ''' </summary>
        Private Delegate Sub TaskDelegate(session As HttpSessionStateBase, data As TaskRequestViewModel)
        ''' <summary>
        ''' Background worker for the task, processes each task and updates progress as it goes
        ''' </summary>
        Private Sub TaskWorker(session As HttpSessionStateBase, data As TaskRequestViewModel)
            'Create a progress model from the request model
            Dim progress As New TaskProgressViewModel()
            For Each item As Task In data.Tasks
                If item.Selected = True Then progress.Tasks.Add(item)
            Next
            progress.Finished = False
            'We are using the session to store our progress data but for a non demonstration solution we would use a database or some other back end store
            Dim repo As New SessionRepository(session)
            If repo.Keys.Contains(ProgressID) Then repo.Remove(ProgressID)
            repo.Add(ProgressID, progress)
            'Process each task updating their status and the repository as we go
            For Each item As Task In progress.Tasks
                With item
                    'Start
                    Thread.Sleep(3000) 'Simulates doing some work
                    .Status = "Started"
                    .Progress = 10
                    'Update the repository with the changes
                    repo.Item(ProgressID) = progress
                    'Step 1
                    Thread.Sleep(2000)
                    .Status = "Step 1 completed"
                    .Progress = 30
                    repo.Item(ProgressID) = progress
                    'Step 2
                    Thread.Sleep(2000)
                    .Status = "Step 2 completed"
                    .Progress = 60
                    repo.Item(ProgressID) = progress
                    'Step 3
                    Thread.Sleep(2000)
                    .Status = "Step 3 completed"
                    .Progress = 90
                    repo.Item(ProgressID) = progress
                    'Finish
                    Thread.Sleep(2000)
                    .Status = "All steps completed"
                    .Progress = 100
                    repo.Item(ProgressID) = progress
                End With
            Next
            progress.Finished = True
            'Update the repository one last time
            repo.Item(ProgressID) = progress
        End Sub
        ''' <summary>
        ''' The taskMonitor does not need any data as it fetches what it requires using Javascript.
        ''' </summary>
        Function TaskMonitor() As ActionResult
            Return View()
        End Function
        ''' <summary>
        ''' The GetTaskProgress function is called to retrieve the current progress of the task.
        ''' </summary>
        Function GetTaskProgress() As ActionResult
            Dim repo As New SessionRepository(Session)
            Dim progress As TaskProgressViewModel = CType(repo.Item(ProgressID), TaskProgressViewModel)
            Return PartialView("_TaskProgress", progress)
        End Function

    End Class

End Namespace