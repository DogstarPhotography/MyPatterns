Imports System.Web.Mvc

Namespace Controllers
    Public Class CustomTemplatesController
        Inherits Controller
        ''' <summary>
        ''' Generate all the data we need here
        ''' </summary>
        Function Index() As ActionResult
            Dim model As New CustomTemplatesViewModel
            model.TableA = New DataListViewModel
            With model.TableA.Items
                .Add(New DataItem With {.Name = "Test_1", .Description = "Test item A"})
                .Add(New DataItem With {.Name = "Test_2", .Description = "Test item B"})
                .Add(New DataItem With {.Name = "Test_3", .Description = "Test item C"})
            End With
            model.TableB = New List(Of DataItem) From {
                New DataItem With {.Name = "Test_4", .Description = "Test item D"},
                New DataItem With {.Name = "Test_5", .Description = "Test item E"},
                New DataItem With {.Name = "Test_6", .Description = "Test item F"}
            }
            Return View(model)
        End Function
        ''' <summary>
        ''' Return a partial view that uses the DataListViewModel.vbhtml template to display the data
        ''' </summary>
        Public Function DataList(data As DataListViewModel) As ActionResult
            Return PartialView("DataList", data)
        End Function
        ''' <summary>
        ''' Return a partial view that uses the DataItem.vbhtml template to display the data
        ''' </summary>
        Public Function DataItems(data As IEnumerable(Of DataItem)) As ActionResult
            Return PartialView("DataItems", data)
        End Function

    End Class
End Namespace