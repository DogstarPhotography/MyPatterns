Imports System.Web.Mvc

Namespace Controllers
    ''' <summary>
    ''' Controller for producing the various partial views used in the layout and other pages
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SharedViewsController
        Inherits Controller

        ''' <summary>
        ''' Return a partial view of the bootstrap navbar. 
        ''' </summary>
        ''' <param name="current">String</param>
        ''' <returns>ActionResult</returns>
        ''' <remarks>Takes a string indicating the current 'page' (i.e. controller) in order to disable an item on the menu.</remarks>
        <ChildActionOnly>
        Public Function NavBar(Optional current As String = "") As ActionResult
            Dim newnavbarmodel As New NavBarViewModel
            newnavbarmodel.Current = current.ToLower
            Return PartialView("_navbar", newnavbarmodel)
        End Function
        ''' <summary>
        ''' Add a button that calls the example modal dialog below.
        ''' </summary>
        <ChildActionOnly>
        Public Function ExampleDialogButton() As ActionResult
            Return PartialView("_ExampleDialogButton")
        End Function
        ''' <summary>
        ''' Add an example modal dialog to the view this is called from.
        ''' </summary>
        <ChildActionOnly>
        Public Function ExampleDialog() As ActionResult
            Return PartialView("_ExampleDialog")
        End Function

    End Class
End Namespace