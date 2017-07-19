Imports Microsoft.Practices.EnterpriseLibrary.Common.Configuration
Imports Microsoft.Practices.EnterpriseLibrary.Data
Imports System.Data.Common
Imports System.Text
Imports Microsoft.Practices.EnterpriseLibrary.Data.Sql
Imports System.ComponentModel

'See http://aspalliance.com/688_Get_Started_with_the_Enterprise_Library_Data_Access_Application_Block for a walkthrough with earlier version of EL
'Search for, download, and install the Microsoft Enterprise Library 5.0 - Hands On Labs
'Then see 'C:\Users\[USER_NAME_HERE]\Documents\Microsoft Enterprise Library 5.0 - Hands On Labs\VB\Data Access\instructions\Data Access instructions VB.xps' for more sample code

Public Class DAAB

	Private db As Database

	Private Sub DAAB_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

		'GetInstance(Of Database)
		Try
			'Create a database from the enterprise library. 'EnterpriseLibrary_Patterns' matches the values in the app.config file:
			'<dataConfiguration defaultDatabase="EnterpriseLibrary_Patterns" />
			'<connectionStrings><add name="EnterpriseLibrary_Patterns" connectionString=...
			db = EnterpriseLibraryContainer.Current.GetInstance(Of Database)("EnterpriseLibrary_Patterns")
			'Alternative option:
			'db = DatabaseFactory.CreateDatabase() 'Uses the value in '<dataConfiguration defaultDatabase=' above
			'db = DatabaseFactory.CreateDatabase("EnterpriseLibrary_Patterns")
			'Dim sqldb As SqlDatabase = EnterpriseLibraryContainer.Current.GetInstance(Of SqlDatabase)()

		Catch ex As Exception
			MessageBox.Show(ex.Message)
		End Try

	End Sub

	Private Sub btnScalar_Click(sender As System.Object, e As System.EventArgs) Handles btnScalar.Click

		'ExecuteScalar
		Dim count As Integer = CType(db.ExecuteScalar(CommandType.Text, "SELECT COUNT(*) FROM SampleData"), Integer)
		Dim message As String = String.Format("There are {0} rows in the sample data table.", count.ToString())
		MessageBox.Show(message)

	End Sub

	Private Sub btnReader_Click(sender As System.Object, e As System.EventArgs) Handles btnReader.Click

		'ExecuteReader
		Using reader As IDataReader = db.ExecuteReader("GetSampleData")
			Dim sb As New StringBuilder
			While reader.Read() = True
				sb.AppendLine("--- New Row ---")
				sb.AppendLine("ID: " & reader.GetInt32(0))
				sb.AppendLine("Name: " & reader.GetString(1))
				sb.AppendLine("Value: " & reader.GetString(2))
			End While
			txtResults.Text = sb.ToString
		End Using

	End Sub

	Private Sub btnDataSetToGrid_Click(sender As System.Object, e As System.EventArgs) Handles btnDataSetToGrid.Click

		'ExecuteDataSet
		Dim ds As DataSet = db.ExecuteDataSet(CommandType.Text, "SELECT * FROM SampleData")
		'Bind the default table in the dataset to the datagridview
		dgvSampleData.DataSource = ds.Tables(0)

	End Sub

	Private Sub btnInsert_Click(sender As System.Object, e As System.EventArgs) Handles btnInsert.Click

		'GetStoredProcCommand
		Dim cmd As DbCommand = db.GetStoredProcCommand("InsertSampleData")
		db.AddInParameter(cmd, "name", DbType.String, "Insert")
		db.AddInParameter(cmd, "value", DbType.String, "New")
		db.ExecuteNonQuery(cmd)

	End Sub

	Private Sub btnSingleRow_Click(sender As System.Object, e As System.EventArgs) Handles btnSingleRow.Click

		'AddInParameter
		Dim cmd As DbCommand = db.GetStoredProcCommand("GetSampleDataItem")
		db.AddInParameter(cmd, "id", DbType.Int32, 1)
		Using reader As IDataReader = db.ExecuteReader(cmd)
			Dim sb As New StringBuilder
			If reader.Read() = True Then
				sb.AppendLine("ID: " & reader.GetInt32(0))
				sb.AppendLine("Name: " & reader.GetString(1))
				sb.AppendLine("Value: " & reader.GetString(2))
			End If
			txtResults.Text = sb.ToString
		End Using

	End Sub

	Private Sub btnOutParameter_Click(sender As System.Object, e As System.EventArgs) Handles btnOutParameter.Click

		'AddOutParameter
		Dim cmd As DbCommand = db.GetStoredProcCommand("GetSampleDataItemNameValue")
		db.AddInParameter(cmd, "id", DbType.Int32, 1)
		db.AddOutParameter(cmd, "name", DbType.String, 50)
		db.AddOutParameter(cmd, "value", DbType.String, -1)	'Note that using -1 returns the correct length
		db.ExecuteNonQuery(cmd)
		txtResults.Text = String.Format("Name: {0}, Value: {1}", cmd.Parameters(0).Value, cmd.Parameters(1).Value)

	End Sub

	Private Sub btnDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnDelete.Click

		'ExecuteNonQuery
		'db.ExecuteNonQuery(CommandType.Text, "DELETE FROM SampleData where [Name] = 'Insert'")

		'GetSqlStringCommand and ExecuteNonQuery
		Dim cmd As DbCommand = db.GetSqlStringCommand("DELETE FROM SampleData where [Name] = 'Insert'")
		db.ExecuteNonQuery(cmd)

	End Sub

	Private Sub btnAccessor_Click(sender As System.Object, e As System.EventArgs) Handles btnAccessor.Click

		'ExecuteSprocAccessor: Executes a stored procedure and returns the result as an enumerable of TResult. 
		'	The conversion from IDataRecord to TResult will be done for each property based on matching property name to column name.
		Dim list = db.ExecuteSprocAccessor(Of SampleData)("GetSampleData")
		Dim results = From s In list Where s.Name = "Test" Order By s.Value Descending
		'Note that we now have a list of objects
		dgvSampleData.DataSource = results.ToArray

	End Sub

	<Description("Fill a DataSet and update the source data")> _
	Private Sub FillAndUpdateDataset()

		Dim selectSQL As String = "SELECT Id, Name, Value FROM SampleData WHERE Id > 0"
		Dim addSQL As String = "INSERT INTO Products (Name, Value) VALUES (@name, @value);"
		Dim updateSQL As String = "UPDATE SampleData SET Name = @name, Value = @value WHERE Id = @id"
		Dim deleteSQL As String = "DELETE FROM SampleData WHERE Id = @id"

		' Fill a DataSet from the Products table using the simple approach
		Dim simpleDS As DataSet = db.ExecuteDataSet(CommandType.Text, selectSQL)
		'DisplayTableNames(simpleDS, "ExecuteDataSet")
		simpleDS = Nothing

		' Fill a DataSet from the Products table using the LoadDataSet method
		' This allows you to specify the name(s) for the table(s) in the DataSet
		Dim loadedDS As New DataSet("ProductsDataSet")
		db.LoadDataSet(CommandType.Text, selectSQL, loadedDS, New String() {"Products"})
		'DisplayTableNames(loadedDS, "LoadDataSet")

		' Update some data in the rows of the DataSet table
		Dim dt As DataTable = loadedDS.Tables("Products")
		dt.Rows(0).Delete()
		Dim rowData As Object() = New Object() {-1, "A New Row", "Added to the table at " & DateTime.Now.ToShortTimeString()}
		dt.Rows.Add(rowData)
		rowData = dt.Rows(1).ItemArray
		rowData(2) = "A new description at " & DateTime.Now.ToShortTimeString()
		dt.Rows(1).ItemArray = rowData
		'DisplayRowValues(dt)

		' Create the commands to update the original table in the database
		Dim insertCommand As DbCommand = db.GetSqlStringCommand(addSQL)
		db.AddInParameter(insertCommand, "name", DbType.[String], "Name", DataRowVersion.Current)
		db.AddInParameter(insertCommand, "description", DbType.[String], "Description", DataRowVersion.Current)
		Dim updateCommand As DbCommand = db.GetSqlStringCommand(updateSQL)
		db.AddInParameter(updateCommand, "name", DbType.[String], "Name", DataRowVersion.Current)
		db.AddInParameter(updateCommand, "description", DbType.[String], "Description", DataRowVersion.Current)
		db.AddInParameter(updateCommand, "id", DbType.[String], "Id", DataRowVersion.Original)
		Dim deleteCommand As DbCommand = db.GetSqlStringCommand(deleteSQL)
		db.AddInParameter(deleteCommand, "id", DbType.Int32, "Id", DataRowVersion.Original)

		' Apply the updates in the DataSet to the original table in the database
		Dim rowsAffected As Integer = db.UpdateDataSet(loadedDS, "Products", insertCommand, updateCommand, deleteCommand, UpdateBehavior.Standard)
		Console.WriteLine("Updated a total of {0} rows in the database.", rowsAffected)

	End Sub

End Class