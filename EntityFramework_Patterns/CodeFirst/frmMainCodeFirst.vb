Imports System.Data.Entity

Public Class frmMainCodeFirst

	Private Sub Form1_Load(sender As Object, e As System.EventArgs) Handles Me.Load

		'Database.SetInitializer(New CodeFirstDataModelContextInitiliazer)

		Try
			'This creates an EntityFramework_CodeFirst.ProductCatalog database when first run!
			Using context As New ProductCatalog
				Dim food = New Category() With {.CategoryId = "FOOD", .Name = "Foods"}
				context.Categories.Add(food) 'This line will fail unless the code first items are moved to a separate project - see CodeFirst.txt
				Dim RecordsAffected As Integer = context.SaveChanges()
				Debug.Print(CStr(RecordsAffected))
				Dim allfoods = (From p In context.Products Where p.CategoryId = "FOOD" Order By p.Name Select p).ToList
				dgvData.DataSource = allfoods
			End Using

		Catch ex As Exception
			Debug.Print(ex.Message)
		End Try

	End Sub
End Class
