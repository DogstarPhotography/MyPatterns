Imports System.Web.Mvc
Imports System.IO

Namespace Controllers
    ''' <summary>
    ''' Demonstrates uploading a file as part of a set of form data and including a progress bar for user feedback.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class XHRUploadController
        Inherits Controller

        Function Index() As ActionResult
            Return View()
        End Function
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="model"></param>
        ''' <returns></returns>
        ''' <remarks>Noite that there is a 30mb file upload limit which is set in the server config files.</remarks>
        <HttpPost>
        Function StoreFormData(model As FileUploadViewModel) As ActionResult
            Dim bytes() As Byte

            ReDim bytes(model.File.ContentLength)
            Dim fs As Stream = model.File.InputStream
            fs.Read(bytes, 0, model.File.ContentLength)
            fs.Flush()
            'Save bytes to back end
            'FURepo.StoreFile(model.File.FileName, bytes)

            Return View("Index")
        End Function

    End Class
End Namespace