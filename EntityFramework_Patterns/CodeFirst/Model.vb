Imports System.Data.Entity

Public Class Category

	Public Property CategoryId As String
	Public Property Name As String

	Public Overridable Property Products As ICollection(Of Product)

End Class

Public Class Product

	Public Property ProductId As Integer
	Public Property Name As String
	Public Property CategoryId As String

	Public Overridable Property Category As Category

End Class

Public Class ProductCatalog
	Inherits DbContext

	Public Property Categories As DbSet(Of Category)
	Public Property Products As DbSet(Of Product)

	Public Sub New()
		MyBase.New("ProductCatalog") 'EntityFramework_CodeFirst.ProductCatalog
	End Sub

End Class