Imports System.Data.Entity
Imports System.Data.Entity.ModelConfiguration.Conventions

Partial Public Class CodeFirstDataModelContext
	Inherits DbContext

	Public Property DataModels As DbSet(Of CodeFirstDataModel)

	Protected Overrides Sub OnModelCreating(modelbuilder As DbModelBuilder)
		'Tell Code First to ignore PluralizingTableName convention
		'If you keep this convention then the generated tables will have pluralized names.
		modelbuilder.Conventions.Remove(Of PluralizingTableNameConvention)()
	End Sub

End Class
