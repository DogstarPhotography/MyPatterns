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

using System.Text;
using Microsoft.Practices.EnterpriseLibrary.Common.Configuration;
using Microsoft.Practices.EnterpriseLibrary.Data;
using System.IO;


/// <summary>
/// Stores content as a file in a database thus allowing files to be stored in a database as if it were a disk drive.
/// </summary>
/// <remarks></remarks>
namespace EnterpriseLibrary_Patterns
{
	public class Datafile
	{
		
#region Constructors
		
		public Datafile()
		{
		}
		
		public Datafile(string category, string group, string filename, string extension)
		{
			this.ID = -1;
			this.Category = category;
			this.Group = group;
			this.Filename = filename;
			this.Extension = DatafilesBackingStore.CleanExtension(extension);
		}
		
		public Datafile(string category, string group, string filename, string extension, byte[] content)
		{
			this.ID = -1;
			this.Category = category;
			this.Group = group;
			this.Filename = filename;
			this.Extension = DatafilesBackingStore.CleanExtension(extension);
			this.Content = content;
			this.Save();
		}
		
		public Datafile(int id, string category, string group, string filename, string extension, byte[] content)
		{
			this.ID = id;
			this.Category = category;
			this.Group = group;
			this.Filename = filename;
			this.Extension = extension;
			this.Content = content;
			this.Save();
		}
		
		public Datafile(string category, string group, string filepath)
		{
			this.ID = -1;
			this.Category = category;
			this.Group = group;
			this.Filename = Path.GetFileNameWithoutExtension(filepath);
			this.Extension = DatafilesBackingStore.CleanExtension(Path.GetExtension(filepath));
			LoadContentFromFile(filepath);
			this.Save();
		}
		
#endregion
		
#region Properties
		/// <summary>
		/// Database unique identifier.
		/// </summary>
		/// <value></value>
		/// <returns></returns>
		/// <remarks></remarks>
		public int ID {get; set;}
		/// <summary>
		/// File group for sorting.
		/// </summary>
		/// <value></value>
		/// <returns></returns>
		/// <remarks></remarks>
		public string Group {get; set;}
		/// <summary>
		/// File category for sorting.
		/// </summary>
		/// <value></value>
		/// <returns></returns>
		/// <remarks></remarks>
		public string Category {get; set;}
		/// <summary>
		/// The file name.
		/// </summary>
		/// <value></value>
		/// <returns></returns>
		/// <remarks></remarks>
		public string Filename {get; set;}
		/// <summary>
		/// The file extension.
		/// </summary>
		/// <value></value>
		/// <returns></returns>
		/// <remarks></remarks>
		public string Extension {get; set;}
		/// <summary>
		/// The actual content of the file.
		/// </summary>
		/// <value></value>
		/// <returns></returns>
		/// <remarks></remarks>
		public byte[] Content {get; set;}
		/// <summary>
		/// Return the file content as text.
		/// </summary>
		/// <value></value>
		/// <returns></returns>
		/// <remarks></remarks>
public string Text
		{
			get
			{
				//Return UTF8Encoding.ASCII.GetString(content) 'Note that this line leaves the Byte Order Mark (BOM) in the returned string
				return Encoding.UTF8.GetString(this.Content); //Whereas this doesn't. Also see http://andrewmatthewthompson.blogspot.co.uk/2011/02/byte-order-mark-found-using-net.html
			}
		}
		/// <summary>
		/// Read a file to load the datafile content property.
		/// </summary>
		/// <param name="filepath"></param>
		/// <remarks></remarks>
		public void LoadContentFromFile(string filepath)
		{
			//Load content from file
			using (FileStream fs = new FileStream(filepath, FileMode.Open, FileAccess.Read))
			{
				using (BinaryReader br = new BinaryReader(fs))
				{
					byte[] data = new byte[(int) fs.Length + 1];
					data = br.ReadBytes((int) fs.Length);
					this.Content = data;
				}
			}
			
			this.Save();
		}
		
#endregion
		
#region CRUD
		/// <summary>
		/// Save this datafile into the database.
		/// </summary>
		/// <remarks></remarks>
		public void Save()
		{
			Datafile retrieved = DatafilesBackingStore.Instance.RetrieveDatafile(this.ID);
			if (retrieved == null)
			{
				this.ID = DatafilesBackingStore.Instance.CreateDatafile(this);
			}
			else
			{
				DatafilesBackingStore.Instance.UpdateDatafile(this);
			}
		}
		/// <summary>
		/// Load a datafile from the database by ID replacing the current contents of this datafile.
		/// </summary>
		/// <param name="id"></param>
		/// <remarks></remarks>
		public void Load(int id)
		{
			this.ID = id;
			this.Load();
		}
		/// <summary>
		/// Load this datafile from the database.
		/// </summary>
		/// <remarks></remarks>
		public void Load()
		{
			Datafile retrieved = DatafilesBackingStore.Instance.RetrieveDatafile(this.ID);
			if (retrieved != null)
			{
				this.Category = retrieved.Category;
				this.Group = retrieved.Group;
				this.Filename = retrieved.Filename;
				this.Extension = retrieved.Extension;
				this.Content = retrieved.Content;
			}
		}
		/// <summary>
		/// Delete this datafile from the database.
		/// </summary>
		/// <remarks></remarks>
		public void Delete()
		{
			DatafilesBackingStore.Instance.DeleteDatafile(this.ID);
			//Me.ID = -1
			//Me.Category = ""
			//Me.Group = ""
			//Me.Filename = ""
			//Me.Extension = ""
			//Me.Content = Nothing
		}
		
#endregion
		
		public void SaveAs(string tempfilename)
		{
			throw (new NotImplementedException());
		}
		
		
	}
	
}
