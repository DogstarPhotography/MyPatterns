Imports Microsoft.Practices.EnterpriseLibrary.Common.Configuration
Imports Microsoft.Practices.EnterpriseLibrary.Data

Public MustInherit Class EnterpriseLibraryDataObjectBase

	Protected Property db As Database
	Protected Property ID As Integer
	''' <summary>
	''' Connect to the database instance.
	''' </summary>
	''' <remarks></remarks>
	Protected Sub ConnectDB()
		Try
			'Create a database from the enterprise library. 'EnterpriseLibrary_Patterns' matches the values in the app.config file:
			'<dataConfiguration defaultDatabase="EnterpriseLibrary_Patterns" />
			'<connectionStrings><add name="EnterpriseLibrary_Patterns" connectionString=...
			db = EnterpriseLibraryContainer.Current.GetInstance(Of Database)("EnterpriseLibrary_Patterns")
		Catch ex As Exception
			Throw ex
		End Try
	End Sub
	''' <summary>
	''' Load this object from the database.
	''' </summary>
	''' <remarks></remarks>
	Public Overridable Sub Load()
		Dim retrieved As EnterpriseLibraryDataObjectBase = Retrieve(Me.ID)
		If retrieved Is Nothing Then
			Me.ID = retrieved.ID
		End If
	End Sub
	''' <summary>
	''' Save this object to the database.
	''' </summary>
	''' <remarks></remarks>
	Public Overridable Sub Save()
		Dim retrieved As EnterpriseLibraryDataObjectBase = Retrieve(Me.ID)
		If retrieved Is Nothing Then
			Me.ID = Create()
		Else
			Update(Me)
		End If
	End Sub
	''' <summary>
	''' Delete this object from the database.
	''' </summary>
	''' <remarks></remarks>
	Public Overridable Sub Delete()
		Delete(Me.ID)
	End Sub

	Protected MustOverride Function Create() As Integer
	Protected MustOverride Function Retrieve(id As Integer) As EnterpriseLibraryDataObjectBase
	Protected MustOverride Sub Update(item As EnterpriseLibraryDataObjectBase)
	Protected MustOverride Sub Delete(id As Integer)

End Class
