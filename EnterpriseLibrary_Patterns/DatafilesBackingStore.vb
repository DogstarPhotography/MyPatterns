Imports Microsoft.Practices.EnterpriseLibrary.Common.Configuration
Imports Microsoft.Practices.EnterpriseLibrary.Data
Imports System.Data.Common
Imports Microsoft.Practices.EnterpriseLibrary.Data.Sql

Public NotInheritable Class DatafilesBackingStore

	Private store As Database

#Region "Singleton"
	Public Shared ReadOnly Instance As DatafilesBackingStore = New DatafilesBackingStore
	''' <summary>
	''' Prevent anyone from instantiating this class by making New private
	''' </summary>
	''' <remarks></remarks>
	Private Sub New()

		'GetInstance(Of Database)
		Try
			'Create a database from the enterprise library. 'EnterpriseLibrary_Patterns' matches the values in the app.config file:
			'<dataConfiguration defaultDatabase="EnterpriseLibrary_Patterns" />
			'<connectionStrings><add name="EnterpriseLibrary_Patterns" connectionString=...
			store = EnterpriseLibraryContainer.Current.GetInstance(Of Database)("EnterpriseLibrary_Patterns")

		Catch ex As Exception
			Throw ex

		End Try

	End Sub

#End Region

#Region "Library functions"
	''' <summary>
	''' Remove the '.' from the extension.
	''' </summary>
	''' <param name="extension"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Function CleanExtension(extension As String) As String
		Dim ext As String = extension
		If ext.Substring(0, 1).Contains(".") Then ext = ext.Substring(1, ext.Length - 1)
		Return ext
	End Function

#End Region

#Region "DataFiles"
	''' <summary>
	''' Retrieve all files for a category and group
	''' </summary>
	''' <param name="category"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Function RetrieveDataFiles(category As String, group As String) As List(Of Datafile)
		Dim files As New List(Of Datafile)
		Dim cmd As DbCommand = store.GetStoredProcCommand("RetrieveDatafiles")
		store.AddInParameter(cmd, "category", DbType.String, category)
		store.AddInParameter(cmd, "group", DbType.String, group)
		Using reader As IDataReader = store.ExecuteReader(cmd)
			While reader.Read
				Dim newfile As Datafile = New Datafile()
				With newfile
					.ID = reader.GetInt32(reader.GetOrdinal("id"))
					.Category = category
                    .Group = reader.SafeGetString(reader.GetOrdinal("group"))
                    .Filename = reader.SafeGetString(reader.GetOrdinal("filename"))
                    .Extension = reader.SafeGetString(reader.GetOrdinal("extension"))
					.Content = CType(reader.GetValue(reader.GetOrdinal("content")), Byte())
				End With
				files.Add(newfile)
			End While
		End Using

	End Function
	''' <summary>
	''' Create a new record in the database.
	''' </summary>
	''' <param name="file"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Function CreateDatafile(file As Datafile) As Integer
		Dim newid As Integer
		Dim cmd As DbCommand = store.GetStoredProcCommand("CreateDatafile")
		store.AddInParameter(cmd, "category", DbType.String, file.Category)
		store.AddInParameter(cmd, "group", DbType.String, file.Group)
		store.AddInParameter(cmd, "filename", DbType.String, file.Filename)
		store.AddInParameter(cmd, "extension", DbType.String, file.Extension)
		store.AddInParameter(cmd, "content", DbType.Binary, file.Content)
		store.AddOutParameter(cmd, "newid", DbType.Int32, newid)
		store.ExecuteNonQuery(cmd)
		newid = CInt(cmd.Parameters(5).Value)
		Return newid
	End Function
	''' <summary>
	''' Retrieve a record from the database.
	''' </summary>
	''' <param name="id"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Function RetrieveDatafile(id As Integer) As Datafile
		Dim cmd As DbCommand = store.GetStoredProcCommand("RetrieveDatafileByID")
		store.AddInParameter(cmd, "id", DbType.String, id)
		Using reader As IDataReader = store.ExecuteReader(cmd)
			If reader.Read() = True Then
				Dim newfile As New Datafile()
				With newfile
					.ID = reader.GetInt32(reader.GetOrdinal("id"))
                    .Category = reader.SafeGetString(reader.GetOrdinal("category"))
                    .Group = reader.SafeGetString(reader.GetOrdinal("group"))
                    .Filename = reader.SafeGetString(reader.GetOrdinal("filename"))
                    .Extension = reader.SafeGetString(reader.GetOrdinal("extension"))
					.Content = CType(reader.GetValue(reader.GetOrdinal("content")), Byte())
				End With
				Return newfile
			Else
				Return Nothing
			End If
		End Using
	End Function
	''' <summary>
	''' Update a record in the database.
	''' </summary>
	''' <param name="file"></param>
	''' <remarks></remarks>
	Public Sub UpdateDatafile(file As Datafile)
		Dim cmd As DbCommand = store.GetStoredProcCommand("UpdateDatafile")
		store.AddInParameter(cmd, "id", DbType.Int32, file.ID)
		store.AddInParameter(cmd, "category", DbType.String, file.Category)
		store.AddInParameter(cmd, "group", DbType.String, file.Group)
		store.AddInParameter(cmd, "filename", DbType.String, file.Filename)
		store.AddInParameter(cmd, "extension", DbType.String, file.Extension)
		store.AddInParameter(cmd, "content", DbType.Binary, file.Content)
		store.ExecuteNonQuery(cmd)
	End Sub
	''' <summary>
	''' Delete a record in the database.
	''' </summary>
	''' <param name="id"></param>
	''' <remarks></remarks>
	Public Sub DeleteDatafile(id As Integer)
		Dim cmd As DbCommand = store.GetStoredProcCommand("DeleteDatafile")
		store.AddInParameter(cmd, "id", DbType.Int32, id)
		store.ExecuteNonQuery(cmd)
	End Sub

#End Region

End Class
