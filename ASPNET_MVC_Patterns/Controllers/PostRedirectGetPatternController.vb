Imports System.Web.Mvc

Namespace Controllers
    ''' <summary>
    ''' Demonstrates the Post-Redirect-Get Pattern:
    ''' <list>
    ''' <item>Handle a POST request from the browser</item>
    ''' <item>Return a Redirect result</item>
    ''' <item>Handle a GET request from the browser</item>
    ''' </list>
    ''' </summary>
    ''' <remarks>
    ''' The advantage of this pattern is that the user does not reuse the POST URL and thus reduces the risk of multiple POSTs
    ''' </remarks>
    Public Class PostRedirectGetPatternController
        Inherits Controller
        ''' <summary>
        ''' Start here
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function PostRedirectGetIndex() As ActionResult
            Dim model As New AjaxFormsViewModel With {.Name = "Test", .Description = "Test description goes here", .Value = 12}
            Return View(model)
        End Function
        ''' <summary>
        ''' Handle user browser POST then redirect
        ''' </summary>
        ''' <param name="model"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function PostData(model As AjaxFormsViewModel) As RedirectToRouteResult
            'Update the model
            Dim updated As New AjaxFormsViewModel
            updated.Name = model.Name & "_updated"
            updated.Description = model.Description & "_updated"
            updated.Value = model.Value * 2
            'Store the updated model to a backing store (we use TempData for demonstration purposes)
            TempData.Add("ViewModel", updated)
            'Redirect to results page
            Return RedirectToAction("PostRedirectGetDisplayData")
        End Function
        ''' <summary>
        ''' Display Results
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function PostRedirectGetDisplayData() As ActionResult
            ' Retrieve the model from the backing store
            Dim model As AjaxFormsViewModel = CType(TempData.Item("ViewModel"), AjaxFormsViewModel)
            'We need to call ModelState.Clear to remove any existing data before we use our own, 
            '   this call clears the data stored in ModelState which the Razor code in the View will use _before_ using the object explicitly passed...!
            ModelState.Clear()
            Return View(model)
        End Function

    End Class
End Namespace