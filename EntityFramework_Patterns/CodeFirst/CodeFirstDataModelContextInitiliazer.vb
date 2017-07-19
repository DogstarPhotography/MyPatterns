Imports System.Data.Entity

Public Class CodeFirstDataModelContextInitiliazer
	Inherits DropCreateDatabaseAlways(Of CodeFirstDataModelContext)

	Protected Overrides Sub Seed(context As CodeFirstDataModelContext)

		Dim model = New CodeFirstDataModel() With {.Index = -1, .Name = "Test", .Value = "Test"}
		context.DataModels.Add(model)
		context.SaveChanges()
		MyBase.Seed(context)

	End Sub
End Class
