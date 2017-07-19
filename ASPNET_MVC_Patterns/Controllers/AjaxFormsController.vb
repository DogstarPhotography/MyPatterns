Imports System.Web.Mvc

Namespace Controllers
    Public Class AjaxFormsController
        Inherits Controller

        Function AjaxFormsIndex() As ActionResult
            Dim model As New AjaxFormsViewModel With {.Name = "Test", .Description = "Test description goes here", .Value = 12}
            Return View(model)
        End Function

#Region "JQuery Ajax Post"

        <HttpPost>
        Function JQueryAjaxPostSaveData(payload As String) As ActionResult
            If payload IsNot Nothing AndAlso payload.Length > 0 Then
                Return New HttpStatusCodeResult(200)
            Else
                Return New HttpStatusCodeResult(500)
            End If
        End Function

        <HttpPost, ValidateInput(False)>
        Function JQueryAjaxPostSaveXML(payload As String) As ActionResult
            If payload IsNot Nothing AndAlso payload.Length > 0 Then
                Return New HttpStatusCodeResult(200)
            Else
                Return New HttpStatusCodeResult(500)
            End If
        End Function

#End Region

#Region "Form Post with Ajax"

        <HttpPost>
        Function FormPostWithAjaxUpdate(model As AjaxFormsViewModel) As JsonResult
            'Update the model
            Dim updated As New AjaxFormsViewModel
            updated.Name = model.Name & "_Updated_via_JSON"
            updated.Description = model.Description & "_Updated_via_JSON"
            updated.Value = model.Value * 10
            ModelState.Clear() 'See above
            'Return as JSON
            Return Json(updated)
        End Function

#End Region

    End Class
End Namespace