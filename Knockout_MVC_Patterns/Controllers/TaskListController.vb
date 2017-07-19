Imports System.Web.Mvc

Namespace Controllers
    ''' <summary>
    ''' Adapted from http://learn.knockoutjs.com/#/?tutorial=loadingsaving
    ''' </summary>
    ''' <remarks></remarks>
    Public Class TaskListController
        Inherits Controller

        Function Index() As ActionResult
            Return View()
        End Function

        <HttpPost>
        Function GetTasks() As JsonResult
            Dim tasks As New List(Of TaskModel)
            tasks.Add(New TaskModel With {.Title = "Wire the money to Panama", .IsDone = True})
            tasks.Add(New TaskModel With {.Title = "Get hair dye, beard trimmer, dark glasses and 'passport'"})
            tasks.Add(New TaskModel With {.Title = "Book taxi to airport"})
            tasks.Add(New TaskModel With {.Title = "Arrange for someone to look after the cat"})
            'This Json function will take our list and turn it into a json object which is then received by an ajax call in the viewmodel
            Return Json(tasks)
        End Function

        <HttpPost>
        Function Save(tasks As IEnumerable(Of TaskModel)) As JsonResult
            'Process data
            Debug.Print(tasks.Count & " tasks")
            For Each task As TaskModel In tasks
                Debug.Print(task.Title & " " & IIf(task.IsDone, "done", "to do").ToString)
            Next
            'This Json function will take our list and turn it into a json object which is then received by an ajax call in the viewmodel
            Return Json(tasks)
        End Function

    End Class
End Namespace