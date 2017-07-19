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


namespace EnterpriseLibrary_Patterns
{
	public abstract class EnterpriseLibraryDataObjectBase
	{
		
		protected Microsoft.Practices.EnterpriseLibrary.Data.Database db {get; set;}
		protected int ID {get; set;}
		/// <summary>
		/// Connect to the database instance.
		/// </summary>
		/// <remarks></remarks>
		protected void ConnectDB()
		{
			try
			{
				//Create a database from the enterprise library. 'EnterpriseLibrary_Patterns' matches the values in the app.config file:
				//<dataConfiguration defaultDatabase="EnterpriseLibrary_Patterns" />
				//<connectionStrings><add name="EnterpriseLibrary_Patterns" connectionString=...
				db = EnterpriseLibraryContainer.Current.GetInstance<Microsoft.Practices.EnterpriseLibrary.Data.Database>("EnterpriseLibrary_Patterns");
			}
			catch (Exception ex)
			{
				throw (ex);
			}
		}
		/// <summary>
		/// Load this object from the database.
		/// </summary>
		/// <remarks></remarks>
		public virtual void Load()
		{
			EnterpriseLibraryDataObjectBase retrieved = Retrieve(this.ID);
			if (retrieved == null)
			{
				this.ID = retrieved.ID;
			}
		}
		/// <summary>
		/// Save this object to the database.
		/// </summary>
		/// <remarks></remarks>
		public virtual void Save()
		{
			EnterpriseLibraryDataObjectBase retrieved = Retrieve(this.ID);
			if (retrieved == null)
			{
				this.ID = Create();
			}
			else
			{
				Update(this);
			}
		}
		/// <summary>
		/// Delete this object from the database.
		/// </summary>
		/// <remarks></remarks>
		public virtual void Delete()
		{
			Delete(this.ID);
		}
		
		protected abstract int Create();
		protected abstract EnterpriseLibraryDataObjectBase Retrieve(int id);
		protected abstract void Update(EnterpriseLibraryDataObjectBase item);
		protected abstract void Delete(int id);
		
	}
	
}
