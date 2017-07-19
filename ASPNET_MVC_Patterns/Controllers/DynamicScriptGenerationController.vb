Imports System.Web.Mvc

Namespace Controllers
    Public Class DynamicScriptGenerationController
        Inherits Controller

        ''' <summary>
        ''' The Index hosts all the generated display and scripts.
        ''' </summary>
        ''' <remarks></remarks>
        Function Index() As ActionResult
            Dim dsgvm As New DynamicScriptGenerationViewModel
            With dsgvm
                .Data.Add(New DynamicScriptGenerationData With {.ID = 1, .Name = "Test A", .Description = "Test data goes here"})
                .Data.Add(New DynamicScriptGenerationData With {.ID = 2, .Name = "Test B", .Description = "Test data goes here"})
                .Data.Add(New DynamicScriptGenerationData With {.ID = 3, .Name = "Test C", .Description = "Test data goes here"})
            End With
            Return View(dsgvm)
        End Function
        ''' <summary>
        ''' This partial view renders the data for display, setting id attributes that will be used by the generated scripts.
        ''' </summary>
        ''' <remarks></remarks>
        Function DisplayData(item As DynamicScriptGenerationData) As PartialViewResult
            Return PartialView(item)
        End Function
        ''' <summary>
        ''' This partial view generates the script needed to manipulate the data
        ''' </summary>
        ''' <remarks></remarks>
        Function GenerateScript(item As DynamicScriptGenerationData) As PartialViewResult
            Return PartialView(item)
        End Function

    End Class
End Namespace