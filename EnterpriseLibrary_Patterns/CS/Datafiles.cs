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

using System.IO;
using Microsoft.Practices.EnterpriseLibrary.Data;
using Microsoft.Practices.EnterpriseLibrary.Common.Configuration;
using System.Data.Common;
using System.Text;


/// <summary>
/// Stores files in a collection with a database backing store, thus effectively treating the DB as a disk drive and the group+category as a directory.
/// </summary>
/// <remarks>
/// TODO: Allow datafiles collection to be archived as a zipped file in the file system using DotNetZip: http://dotnetzip.codeplex.com/
/// TOOD: Allow for scalability of multiple users with multiple archives.
/// TODO: Look into use of SQL Server FILESTREAM to optimise file storage, see http://msdn.microsoft.com/en-us/library/bb933993.aspx
/// </remarks>
namespace EnterpriseLibrary_Patterns
{
	public class Datafiles : System.Collections.Generic.Dictionary<int, Datafile>
	{
		
#region Properties
		
		private Microsoft.Practices.EnterpriseLibrary.Data.Database db;
		private string _category;
		private string _group;
		/// <summary>
		/// Category under which files will be stored, set via constructor.
		/// </summary>
		/// <value></value>
		/// <returns></returns>
		/// <remarks></remarks>
public string Category
		{
			get
			{
				return _category;
			}
		}
		/// <summary>
		/// Current/default group for files added to the FileStore if a group is not explicitly specified.
		/// </summary>
		/// <value></value>
		/// <returns></returns>
		/// <remarks></remarks>
public string Group
		{
			get
			{
				return _group;
			}
			set
			{
				_group = value;
			}
		}
		
#endregion
		
#region Constructor
		/// <summary>
		/// This constructor will set the Category and Group to 'Default'
		/// </summary>
		/// <remarks></remarks>
		public Datafiles()
		{
			_category = "Default";
			this.Group = "Default";
		}
		/// <summary>
		/// This constructor will set the Group to 'Default'
		/// </summary>
		/// <param name="category"></param>
		/// <remarks></remarks>
		public Datafiles(string category)
		{
			_category = category;
			this.Group = "Default";
		}
		/// <summary>
		/// This constructor allows Category and Group to be set.
		/// </summary>
		/// <param name="category"></param>
		/// <param name="group"></param>
		/// <remarks></remarks>
		public Datafiles(string category, string group)
		{
			_category = category;
			this.Group = group;
		}
		
#endregion
		
#region Backing store access
		/// <summary>
		/// Loads the datafiles from the backing store.
		/// </summary>
		/// <remarks></remarks>
		public void LoadDatafiles()
		{
			List<Datafile> files = DatafilesBackingStore.Instance.RetrieveDataFiles(this.Category, this.Group);
			foreach (Datafile file in files)
			{
				this.Add(file.ID, file);
			}
		}
		/// <summary>
		/// Save the datafiles to the backing store
		/// </summary>
		/// <remarks></remarks>
		public void SaveDatafiles()
		{
			foreach (Datafile file in this.Values)
			{
				file.Save();
			}
		}
		/// <summary>
		/// Deletes all files in the backing store and clears the Datafiles collection.
		/// </summary>
		/// <remarks></remarks>
		public void DeleteDatafiles()
		{
			foreach (Datafile file in this.Values)
			{
				file.Delete();
			}
			this.Clear();
		}
		
#endregion
		
#region Add File Helpers
		/// <summary>
		/// Add a file to the collection using the default group.
		/// </summary>
		/// <param name="filepath"></param>
		/// <remarks></remarks>
		public void AddFile(string filepath)
		{
			this.AddFile(this.Group, filepath);
		}
		/// <summary>
		/// Add a file to the collection using the given group.
		/// </summary>
		/// <param name="group"></param>
		/// <param name="filepath"></param>
		/// <remarks></remarks>
		public void AddFile(string group, string filepath)
		{
			//Create new datafile
			Datafile newFile = new Datafile(this.Category, group, filepath);
			//Add to collection
			this.Add(newFile.ID, newFile);
		}
		
#endregion
		
#region Retrieve File Helpers
		/// <summary>
		///  Get the Id of the file with the given parameters in the current/default Group.
		/// </summary>
		/// <param name="filename"></param>
		/// <param name="extension"></param>
		/// <returns></returns>
		/// <remarks></remarks>
		public int GetFileID(string filename, string extension)
		{
			if (extension.Substring(0, 1).Contains("."))
			{
				extension = extension.Substring(1, extension.Length - 1);
			}
			return (from f in this.Values where f.Group == this.Group && f.Filename == filename && f.Extension == extension select f).First().ID;
		}
		/// <summary>
		/// Get the Id of the file with the given parameters.
		/// </summary>
		/// <param name="group"></param>
		/// <param name="filename"></param>
		/// <param name="extension"></param>
		/// <returns></returns>
		/// <remarks></remarks>
		public int GetFileID(string group, string filename, string extension)
		{
			if (extension.Substring(0, 1).Contains("."))
			{
				extension = extension.Substring(1, extension.Length - 1);
			}
			return (from f in this.Values where f.group == group && f.filename == filename && f.extension == extension select f).First.ID;
		}
		/// <summary>
		/// Get a file as binary.
		/// </summary>
		/// <param name="id"></param>
		/// <returns></returns>
		/// <remarks></remarks>
		public byte[] GetBinary(int id)
		{
			return this[id].Content;
		}
		/// <summary>
		/// Get the text for a datafile.
		/// </summary>
		/// <param name="id"></param>
		/// <returns></returns>
		/// <remarks></remarks>
		public string GetText(int id)
		{
			byte[] content = null;
			content = this[id].Content;
			//Return UTF8Encoding.ASCII.GetString(content) 'Note that this line leaves the BOM in the returned string
			return Encoding.UTF8.GetString(content); //Whereas this doesn't. Also see http://andrewmatthewthompson.blogspot.co.uk/2011/02/byte-order-mark-found-using-net.html
		}
		
#endregion
		
#region Archive functions
		/// <summary>
		/// Write all the files to an archive with the given filename and path.
		/// </summary>
		/// <param name="filename"></param>
		/// <remarks></remarks>
		public void Archive(string filename)
		{
			
			//Dim temp As String = Path.GetTempPath()
			//Using zip As ZipFile = New ZipFile
			//	For Each file As Datafile In Me.Values
			//		'Save each file to the temporary directory
			//		Dim tempfilename As String = Path.Combine(temp, file.Filename)
			//		file.SaveAs(tempfilename)
			//		'Add it to the zip
			//		zip.AddFile(tempfilename)
			//		'Delete temp
			//		IO.File.Delete(tempfilename)
			//	Next
			//	zip.Save(filename)
			//End Using
			
		}
		/// <summary>
		/// Load this class from the given filename and path.
		/// </summary>
		/// <param name="filename"></param>
		/// <remarks></remarks>
		public void Unarchive(string filename)
		{
			
		}
		
#endregion
		
	}
	
}
