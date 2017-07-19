Imports System.Text
Imports Microsoft.Practices.EnterpriseLibrary.Common.Configuration
Imports Microsoft.Practices.EnterpriseLibrary.Data
Imports System.IO

''' <summary>
''' Stores content as a file in a database thus allowing files to be stored in a database as if it were a disk drive.
''' </summary>
''' <remarks></remarks>
Public Class Datafile

#Region "Constructors"

	Public Sub New()
	End Sub

	Public Sub New(category As String, group As String, filename As String, extension As String)
		Me.ID = -1
		Me.Category = category
		Me.Group = group
		Me.Filename = filename
		Me.Extension = DatafilesBackingStore.CleanExtension(extension)
	End Sub

	Public Sub New(category As String, group As String, filename As String, extension As String, content As Byte())
		Me.ID = -1
		Me.Category = category
		Me.Group = group
		Me.Filename = filename
		Me.Extension = DatafilesBackingStore.CleanExtension(extension)
		Me.Content = content
		Me.Save()
	End Sub

	Public Sub New(id As Integer, category As String, group As String, filename As String, extension As String, content As Byte())
		Me.ID = id
		Me.Category = category
		Me.Group = group
		Me.Filename = filename
		Me.Extension = extension
		Me.Content = content
		Me.Save()
	End Sub

	Public Sub New(category As String, group As String, filepath As String)
		Me.ID = -1
		Me.Category = category
		Me.Group = group
		Me.Filename = Path.GetFileNameWithoutExtension(filepath)
		Me.Extension = DataFilesBackingStore.CleanExtension(Path.GetExtension(filepath))
		LoadContentFromFile(filepath)
		Me.Save()
	End Sub

#End Region

#Region "Properties"
	''' <summary>
	''' Database unique identifier.
	''' </summary>
	''' <value></value>
	''' <returns></returns>
	''' <remarks></remarks>
	Property ID As Integer
	''' <summary>
	''' File group for sorting.
	''' </summary>
	''' <value></value>
	''' <returns></returns>
	''' <remarks></remarks>
	Property Group As String
	''' <summary>
	''' File category for sorting.
	''' </summary>
	''' <value></value>
	''' <returns></returns>
	''' <remarks></remarks>
	Property Category As String
	''' <summary>
	''' The file name.
	''' </summary>
	''' <value></value>
	''' <returns></returns>
	''' <remarks></remarks>
	Property Filename As String
	''' <summary>
	''' The file extension.
	''' </summary>
	''' <value></value>
	''' <returns></returns>
	''' <remarks></remarks>
	Property Extension As String
	''' <summary>
	''' The actual content of the file.
	''' </summary>
	''' <value></value>
	''' <returns></returns>
	''' <remarks></remarks>
	Property Content As Byte()
	''' <summary>
	''' Return the file content as text.
	''' </summary>
	''' <value></value>
	''' <returns></returns>
	''' <remarks></remarks>
	Public ReadOnly Property Text() As String
		Get
			'Return UTF8Encoding.ASCII.GetString(content) 'Note that this line leaves the Byte Order Mark (BOM) in the returned string
			Return Encoding.UTF8.GetString(Me.Content)	'Whereas this doesn't. Also see http://andrewmatthewthompson.blogspot.co.uk/2011/02/byte-order-mark-found-using-net.html
		End Get
	End Property
	''' <summary>
	''' Read a file to load the datafile content property.
	''' </summary>
	''' <param name="filepath"></param>
	''' <remarks></remarks>
	Public Sub LoadContentFromFile(filepath As String)
		'Load content from file
		Using fs As New FileStream(filepath, FileMode.Open, FileAccess.Read), br As New BinaryReader(fs)
			Dim data(CInt(fs.Length)) As Byte
			data = br.ReadBytes(CInt(fs.Length))
			Me.Content = data
		End Using
		Me.Save()
	End Sub

#End Region

#Region "CRUD"
	''' <summary>
	''' Save this datafile into the database.
	''' </summary>
	''' <remarks></remarks>
	Public Sub Save()
		Dim retrieved As Datafile = DatafilesBackingStore.Instance.RetrieveDatafile(Me.ID)
		If retrieved Is Nothing Then
			Me.ID = DatafilesBackingStore.Instance.CreateDatafile(Me)
		Else
			DatafilesBackingStore.Instance.UpdateDatafile(Me)
		End If
	End Sub
	''' <summary>
	''' Load a datafile from the database by ID replacing the current contents of this datafile.
	''' </summary>
	''' <param name="id"></param>
	''' <remarks></remarks>
	Public Sub Load(id As Integer)
		Me.ID = id
		Me.Load()
	End Sub
	''' <summary>
	''' Load this datafile from the database.
	''' </summary>
	''' <remarks></remarks>
	Public Sub Load()
		Dim retrieved As Datafile = DatafilesBackingStore.Instance.RetrieveDatafile(Me.ID)
		If retrieved IsNot Nothing Then
			Me.Category = retrieved.Category
			Me.Group = retrieved.Group
			Me.Filename = retrieved.Filename
			Me.Extension = retrieved.Extension
			Me.Content = retrieved.Content
		End If
	End Sub
	''' <summary>
	''' Delete this datafile from the database.
	''' </summary>
	''' <remarks></remarks>
	Public Sub Delete()
		DatafilesBackingStore.Instance.DeleteDatafile(Me.ID)
		'Me.ID = -1
		'Me.Category = ""
		'Me.Group = ""
		'Me.Filename = ""
		'Me.Extension = ""
		'Me.Content = Nothing
	End Sub

#End Region

	Sub SaveAs(tempfilename As String)
		Throw New NotImplementedException
	End Sub


End Class
