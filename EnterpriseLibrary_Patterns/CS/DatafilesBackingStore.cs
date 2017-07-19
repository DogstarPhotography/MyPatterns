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
using Microsoft.Practices.EnterpriseLibrary.Data.Sql;


namespace EnterpriseLibrary_Patterns
{
	public sealed class DatafilesBackingStore
	{
		
		private Microsoft.Practices.EnterpriseLibrary.Data.Database store;
		
#region Singleton
		public static readonly DatafilesBackingStore Instance = new DatafilesBackingStore();
		/// <summary>
		/// Prevent anyone from instantiating this class by making New private
		/// </summary>
		/// <remarks></remarks>
		private DatafilesBackingStore()
		{
			
			//GetInstance(Of Database)
			try
			{
				//Create a database from the enterprise library. 'EnterpriseLibrary_Patterns' matches the values in the app.config file:
				//<dataConfiguration defaultDatabase="EnterpriseLibrary_Patterns" />
				//<connectionStrings><add name="EnterpriseLibrary_Patterns" connectionString=...
				store = EnterpriseLibraryContainer.Current.GetInstance<Microsoft.Practices.EnterpriseLibrary.Data.Database>("EnterpriseLibrary_Patterns");
				
			}
			catch (Exception ex)
			{
				throw (ex);
				
			}
			
		}
		
#endregion
		
#region Library functions
		/// <summary>
		/// Remove the '.' from the extension.
		/// </summary>
		/// <param name="extension"></param>
		/// <returns></returns>
		/// <remarks></remarks>
		public static string CleanExtension(string extension)
		{
			string ext = extension;
			if (ext.Substring(0, 1).Contains("."))
			{
				ext = ext.Substring(1, ext.Length - 1);
			}
			return ext;
		}
		
#endregion
		
#region DataFiles
		/// <summary>
		/// Retrieve all files for a category and group
		/// </summary>
		/// <param name="category"></param>
		/// <returns></returns>
		/// <remarks></remarks>
		public List<Datafile> RetrieveDataFiles(string category, string group)
		{
			List<Datafile> files = new List<Datafile>();
			DbCommand cmd = store.GetStoredProcCommand("RetrieveDatafiles");
			store.AddInParameter(cmd, "category", DbType.String, category);
			store.AddInParameter(cmd, "group", DbType.String, group);
			using (IDataReader reader = store.ExecuteReader(cmd))
			{
				while (reader.Read())
				{
					Datafile newfile = new Datafile();
					newfile.ID = reader.GetInt32(reader.GetOrdinal("id"));
					newfile.Category = category;
					newfile.Group = System.Convert.ToString(reader.SafeGetString(reader.GetOrdinal("group")));
					newfile.Filename = System.Convert.ToString(reader.SafeGetString(reader.GetOrdinal("filename")));
					newfile.Extension = System.Convert.ToString(reader.SafeGetString(reader.GetOrdinal("extension")));
					newfile.Content = (byte[]) (reader.GetValue(reader.GetOrdinal("content")));
					files.Add(newfile);
				}
			}
			
			
			return default(List<Datafile>);
		}
		/// <summary>
		/// Create a new record in the database.
		/// </summary>
		/// <param name="file"></param>
		/// <returns></returns>
		/// <remarks></remarks>
		public int CreateDatafile(Datafile file)
		{
			int newid = 0;
			DbCommand cmd = store.GetStoredProcCommand("CreateDatafile");
			store.AddInParameter(cmd, "category", DbType.String, file.Category);
			store.AddInParameter(cmd, "group", DbType.String, file.Group);
			store.AddInParameter(cmd, "filename", DbType.String, file.Filename);
			store.AddInParameter(cmd, "extension", DbType.String, file.Extension);
			store.AddInParameter(cmd, "content", DbType.Binary, file.Content);
			store.AddOutParameter(cmd, "newid", DbType.Int32, newid);
			store.ExecuteNonQuery(cmd);
			newid = System.Convert.ToInt32(cmd.Parameters[5].Value);
			return newid;
		}
		/// <summary>
		/// Retrieve a record from the database.
		/// </summary>
		/// <param name="id"></param>
		/// <returns></returns>
		/// <remarks></remarks>
		public Datafile RetrieveDatafile(int id)
		{
			DbCommand cmd = store.GetStoredProcCommand("RetrieveDatafileByID");
			store.AddInParameter(cmd, "id", DbType.String, id);
			using (System.Data.IDataReader reader = store.ExecuteReader(cmd))
			{
				if (reader.Read() == true)
				{
					Datafile newfile = new Datafile();
					newfile.ID = reader.GetInt32(reader.GetOrdinal("id"));
					newfile.Category = System.Convert.ToString(reader.SafeGetString(reader.GetOrdinal("category")));
					newfile.Group = System.Convert.ToString(reader.SafeGetString(reader.GetOrdinal("group")));
					newfile.Filename = System.Convert.ToString(reader.SafeGetString(reader.GetOrdinal("filename")));
					newfile.Extension = System.Convert.ToString(reader.SafeGetString(reader.GetOrdinal("extension")));
					newfile.Content = (byte[]) (reader.GetValue(reader.GetOrdinal("content")));
					return newfile;
				}
				else
				{
					return null;
				}
			}
			
		}
		/// <summary>
		/// Update a record in the database.
		/// </summary>
		/// <param name="file"></param>
		/// <remarks></remarks>
		public void UpdateDatafile(Datafile file)
		{
			DbCommand cmd = store.GetStoredProcCommand("UpdateDatafile");
			store.AddInParameter(cmd, "id", DbType.Int32, file.ID);
			store.AddInParameter(cmd, "category", DbType.String, file.Category);
			store.AddInParameter(cmd, "group", DbType.String, file.Group);
			store.AddInParameter(cmd, "filename", DbType.String, file.Filename);
			store.AddInParameter(cmd, "extension", DbType.String, file.Extension);
			store.AddInParameter(cmd, "content", DbType.Binary, file.Content);
			store.ExecuteNonQuery(cmd);
		}
		/// <summary>
		/// Delete a record in the database.
		/// </summary>
		/// <param name="id"></param>
		/// <remarks></remarks>
		public void DeleteDatafile(int id)
		{
			DbCommand cmd = store.GetStoredProcCommand("DeleteDatafile");
			store.AddInParameter(cmd, "id", DbType.Int32, id);
			store.ExecuteNonQuery(cmd);
		}
		
#endregion
		
	}
	
}
