Imports System.Web.Mvc

Namespace Controllers
    Public Class LibraryController
        Inherits Controller

        Private Shared library As New LibraryModel

#Region "ASP.NET"

        Function Index() As ActionResult
            Return View(library)
        End Function

        <HttpGet>
        Public Function Edit(id As Integer) As ActionResult
            Return View(library.GetBook(id))
        End Function

        <HttpPost>
        Public Function Edit(book As BookModel) As ActionResult
            library.UpdateBook(book)
            Return RedirectToAction("Index")
        End Function

        Public Function Add() As ActionResult
            Dim book As New BookModel With {
                .Title = "New Book",
                .Author = "Unknown",
                .Year = DateTime.Now.Year
            }
            library.AddBook(book)
            Return RedirectToAction("Index")
        End Function

        Public Function Remove(id As Integer) As ActionResult
            library.RemoveBook(id)
            Return RedirectToAction("Index")
        End Function

#End Region

#Region "Knockout"

        Public Function koGetName() As JsonResult
            Return Json(library.Name, JsonRequestBehavior.AllowGet)
        End Function

        Public Function koGetBooks() As JsonResult
            Return Json(library.GetBooks, JsonRequestBehavior.AllowGet)
        End Function

        Public Function koAdd() As JsonResult
            Dim book As New BookModel With {
                .Title = "New Book",
                .Author = "Unknown",
                .Year = DateTime.Now.Year
                }
            library.AddBook(book)
            Return Json(library.GetBooks, JsonRequestBehavior.AllowGet)
        End Function

        Public Function koRemove(id As Integer) As JsonResult
            library.RemoveBook(id)
            Return Json(library.GetBooks, JsonRequestBehavior.AllowGet)
        End Function

#End Region

    End Class
End Namespace