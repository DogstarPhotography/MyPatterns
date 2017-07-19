Imports System.Web.Mvc
Imports System.IO

'Reference: http://www.prideparrot.com/blog/archive/2012/8/uploading_and_returning_files

Namespace Controllers
    ''' <summary>
    ''' Demonstrates simple file uploading with no progress.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class FileUploadController
        Inherits Controller

        Function Index() As ActionResult
            Return View()
        End Function

        <HttpPost>
        Function StoreFile() As ActionResult
            Dim bytes() As Byte
            For Each item As String In Request.Files
                If Request.Files(item) IsNot Nothing AndAlso Request.Files(item).ContentLength > 0 Then
                    ReDim bytes(Request.Files(item).ContentLength)
                    Using fs As Stream = Request.Files(item).InputStream
                        fs.Read(bytes, 0, Request.Files(item).ContentLength)
                        fs.Flush()
                    End Using
                    'Save bytes to back end
                    'Repository.StoreFile(Request.Files(item).FileName, bytes)
                    Exit For 'We only ever deal with one file
                End If
            Next
            Return View("Index")
        End Function

        <HttpPost>
        Function StoreFormData(model As FileUploadViewModel) As ActionResult
            Dim bytes() As Byte
            ReDim bytes(model.File.ContentLength)
            Using fs As Stream = model.File.InputStream
                fs.Read(bytes, 0, model.File.ContentLength)
                fs.Flush()
            End Using
            'Save bytes to back end
            'Repository.StoreFile(model.File.FileName, model.Name, model.Description, bytes)
            Return View("Index")
        End Function

    End Class
End Namespace