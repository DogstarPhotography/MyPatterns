// VBConversions Note: VB project level imports
using System.Collections.Generic;
using System;
using System.Linq;
using System.Drawing;
using System.Diagnostics;
using System.Data;
using System.Xml.Linq;
using Microsoft.VisualBasic;
using System.Collections;
using System.Windows.Forms;
// End of VB project level imports

using Microsoft.Practices.EnterpriseLibrary.Common.Configuration;
using Microsoft.Practices.EnterpriseLibrary.Data;
using System.Data.Common;
using System.Text;
using Microsoft.Practices.EnterpriseLibrary.Data.Sql;
using System.ComponentModel;


//See http://aspalliance.com/688_Get_Started_with_the_Enterprise_Library_Data_Access_Application_Block for a walkthrough with earlier version of EL
//Search for, download, and install the Microsoft Enterprise Library 5.0 - Hands On Labs
//Then see 'C:\Users\[USER_NAME_HERE]\Documents\Microsoft Enterprise Library 5.0 - Hands On Labs\VB\Data Access\instructions\Data Access instructions VB.xps' for more sample code

namespace EnterpriseLibrary_Patterns
{
	public partial class DAAB
	{
		public DAAB()
		{
			InitializeComponent();
			
			//Added to support default instance behavour in C#
			if (defaultInstance == null)
				defaultInstance = this;
		}
		
#region Default Instance
		
		private static DAAB defaultInstance;
		
		/// <summary>
		/// Added by the VB.Net to C# Converter to support default instance behavour in C#
		/// </summary>
public static DAAB Default
		{
			get
			{
				if (defaultInstance == null)
				{
					defaultInstance = new DAAB();
					defaultInstance.FormClosed += new FormClosedEventHandler(defaultInstance_FormClosed);
				}
				
				return defaultInstance;
			}
		}
		
		static void defaultInstance_FormClosed(object sender, FormClosedEventArgs e)
		{
			defaultInstance = null;
		}
		
#endregion
		
		private Microsoft.Practices.EnterpriseLibrary.Data.Database db;
		
		public void DAAB_Load(System.Object sender, System.EventArgs e)
		{
			
			//GetInstance(Of Database)
			try
			{
				//Create a database from the enterprise library. 'EnterpriseLibrary_Patterns' matches the values in the app.config file:
				//<dataConfiguration defaultDatabase="EnterpriseLibrary_Patterns" />
				//<connectionStrings><add name="EnterpriseLibrary_Patterns" connectionString=...
				db = EnterpriseLibraryContainer.Current.GetInstance<Microsoft.Practices.EnterpriseLibrary.Data.Database>("EnterpriseLibrary_Patterns");
				//Alternative option:
				//db = DatabaseFactory.CreateDatabase() 'Uses the value in '<dataConfiguration defaultDatabase=' above
				//db = DatabaseFactory.CreateDatabase("EnterpriseLibrary_Patterns")
				//Dim sqldb As SqlDatabase = EnterpriseLibraryContainer.Current.GetInstance(Of SqlDatabase)()
				
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
			
		}
		
		public void btnScalar_Click(System.Object sender, System.EventArgs e)
		{
			
			//ExecuteScalar
			int count = System.Convert.ToInt32(db.ExecuteScalar(CommandType.Text, "SELECT COUNT(*) FROM SampleData"));
			string message = string.Format("There are {0} rows in the sample data table.", count.ToString());
			MessageBox.Show(message);
			
		}
		
		public void btnReader_Click(System.Object sender, System.EventArgs e)
		{
			
			//ExecuteReader
			using (IDataReader reader = db.ExecuteReader("GetSampleData"))
			{
				StringBuilder sb = new StringBuilder();
				while (reader.Read() == true)
				{
					sb.AppendLine("--- New Row ---");
sb.AppendLine("ID: " + System.Convert.ToString(reader.GetInt32(0)));
sb.AppendLine("Name: " + reader.GetString(1));
sb.AppendLine("Value: " + reader.GetString(2));
				}
				txtResults.Text = sb.ToString();
			}
			
			
		}
		
		public void btnDataSetToGrid_Click(System.Object sender, System.EventArgs e)
		{
			
			//ExecuteDataSet
			DataSet ds = db.ExecuteDataSet(CommandType.Text, "SELECT * FROM SampleData");
			//Bind the default table in the dataset to the datagridview
			dgvSampleData.DataSource = ds.Tables[0];
			
		}
		
		public void btnInsert_Click(System.Object sender, System.EventArgs e)
		{
			
			//GetStoredProcCommand
			DbCommand cmd = db.GetStoredProcCommand("InsertSampleData");
			db.AddInParameter(cmd, "name", DbType.String, "Insert");
			db.AddInParameter(cmd, "value", DbType.String, "New");
			db.ExecuteNonQuery(cmd);
			
		}
		
		public void btnSingleRow_Click(System.Object sender, System.EventArgs e)
		{
			
			//AddInParameter
			DbCommand cmd = db.GetStoredProcCommand("GetSampleDataItem");
			db.AddInParameter(cmd, "id", DbType.Int32, 1);
			using (IDataReader reader = db.ExecuteReader(cmd))
			{
				StringBuilder sb = new StringBuilder();
				if (reader.Read() == true)
				{
sb.AppendLine("ID: " + System.Convert.ToString(reader.GetInt32(0)));
sb.AppendLine("Name: " + reader.GetString(1));
sb.AppendLine("Value: " + reader.GetString(2));
				}
				txtResults.Text = sb.ToString();
			}
			
			
		}
		
		public void btnOutParameter_Click(System.Object sender, System.EventArgs e)
		{
			
			//AddOutParameter
			DbCommand cmd = db.GetStoredProcCommand("GetSampleDataItemNameValue");
			db.AddInParameter(cmd, "id", DbType.Int32, 1);
			db.AddOutParameter(cmd, "name", DbType.String, 50);
			db.AddOutParameter(cmd, "value", DbType.String, -1); //Note that using -1 returns the correct length
			db.ExecuteNonQuery(cmd);
			txtResults.Text = string.Format("Name: {0}, Value: {1}", cmd.Parameters[0].Value, cmd.Parameters[1].Value);
			
		}
		
		public void btnDelete_Click(System.Object sender, System.EventArgs e)
		{
			
			//ExecuteNonQuery
			//db.ExecuteNonQuery(CommandType.Text, "DELETE FROM SampleData where [Name] = 'Insert'")
			
			//GetSqlStringCommand and ExecuteNonQuery
			DbCommand cmd = db.GetSqlStringCommand("DELETE FROM SampleData where [Name] = \'Insert\'");
			db.ExecuteNonQuery(cmd);
			
		}
		
		public void btnAccessor_Click(System.Object sender, System.EventArgs e)
		{
			
			//ExecuteSprocAccessor: Executes a stored procedure and returns the result as an enumerable of TResult.
			//	The conversion from IDataRecord to TResult will be done for each property based on matching property name to column name.
			var list = db.ExecuteSprocAccessor<SampleData>("GetSampleData");
			var results = from s in list where s.Name == "Test" orderby s.Value descending select s;
			//Note that we now have a list of objects
			dgvSampleData.DataSource = results.ToArray();
			
		}
		
		[Description("Fill a DataSet and update the source data")]private void FillAndUpdateDataset()
		{
			
			string selectSQL = "SELECT Id, Name, Value FROM SampleData WHERE Id > 0";
			string addSQL = "INSERT INTO Products (Name, Value) VALUES (@name, @value);";
			string updateSQL = "UPDATE SampleData SET Name = @name, Value = @value WHERE Id = @id";
			string deleteSQL = "DELETE FROM SampleData WHERE Id = @id";
			
			// Fill a DataSet from the Products table using the simple approach
			DataSet simpleDS = db.ExecuteDataSet(CommandType.Text, selectSQL);
			//DisplayTableNames(simpleDS, "ExecuteDataSet")
			simpleDS = null;
			
			// Fill a DataSet from the Products table using the LoadDataSet method
			// This allows you to specify the name(s) for the table(s) in the DataSet
			DataSet loadedDS = new DataSet("ProductsDataSet");
			db.LoadDataSet(CommandType.Text, selectSQL, loadedDS, new string[] {"Products"});
			//DisplayTableNames(loadedDS, "LoadDataSet")
			
			// Update some data in the rows of the DataSet table
			DataTable dt = loadedDS.Tables["Products"];
			dt.Rows[0].Delete();
			object[] rowData = new object[] {-1, "A New Row", "Added to the table at " + DateTime.Now.ToShortTimeString()};
			dt.Rows.Add(rowData);
			rowData = dt.Rows[1].ItemArray;
			rowData[2] = "A new description at " + DateTime.Now.ToShortTimeString();
			dt.Rows[1].ItemArray = rowData;
			//DisplayRowValues(dt)
			
			// Create the commands to update the original table in the database
			DbCommand insertCommand = db.GetSqlStringCommand(addSQL);
			db.AddInParameter(insertCommand, "name", DbType.String, "Name", DataRowVersion.Current);
			db.AddInParameter(insertCommand, "description", DbType.String, "Description", DataRowVersion.Current);
			DbCommand updateCommand = db.GetSqlStringCommand(updateSQL);
			db.AddInParameter(updateCommand, "name", DbType.String, "Name", DataRowVersion.Current);
			db.AddInParameter(updateCommand, "description", DbType.String, "Description", DataRowVersion.Current);
			db.AddInParameter(updateCommand, "id", DbType.String, "Id", DataRowVersion.Original);
			DbCommand deleteCommand = db.GetSqlStringCommand(deleteSQL);
			db.AddInParameter(deleteCommand, "id", DbType.Int32, "Id", DataRowVersion.Original);
			
			// Apply the updates in the DataSet to the original table in the database
			int rowsAffected = db.UpdateDataSet(loadedDS, "Products", insertCommand, updateCommand, deleteCommand, UpdateBehavior.Standard);
			Console.WriteLine("Updated a total of {0} rows in the database.", rowsAffected);
			
		}
		
	}
}
