Public Class Implementation

    Sub Test()

        Dim Test As String = Config.Settings("Test")

        Dim newWorker As GenericWorker(Of GenericJob(Of GenericTask(Of GenericState(Of Object), Object), GenericState(Of Object), Object), GenericTask(Of GenericState(Of Object), Object), GenericState(Of Object), Object) = New GenericWorker(Of GenericJob(Of GenericTask(Of GenericState(Of Object), Object), GenericState(Of Object), Object), GenericTask(Of GenericState(Of Object), Object), GenericState(Of Object), Object)
        Dim newJob As GenericJob(Of GenericTask(Of GenericState(Of Object), Object), GenericState(Of Object), Object) = New GenericJob(Of GenericTask(Of GenericState(Of Object), Object), GenericState(Of Object), Object)
        Dim newTask As GenericTask(Of GenericState(Of Object), Object) = New GenericTask(Of GenericState(Of Object), Object)
        Dim newState As GenericState(Of Object) = New GenericState(Of Object)

        newTask.State = newState
        newJob.Tasks.Enqueue(newTask)
        newWorker.Jobs.Add(newJob)

        newWorker.StartWork()

    End Sub


End Class
