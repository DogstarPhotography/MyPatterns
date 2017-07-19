Imports System.IO
Imports Microsoft.Practices.EnterpriseLibrary.Data
Imports Microsoft.Practices.EnterpriseLibrary.Common.Configuration
Imports System.Data.Common
Imports System.Text

''' <summary>
''' Stores files in a collection with a database backing store, thus effectively treating the DB as a disk drive and the group+category as a directory.
''' </summary>
''' <remarks>
''' TODO: Allow datafiles collection to be archived as a zipped file in the file system using DotNetZip: http://dotnetzip.codeplex.com/
''' TOOD: Allow for scalability of multiple users with multiple archives.
''' TODO: Look into use of SQL Server FILESTREAM to optimise file storage, see http://msdn.microsoft.com/en-us/library/bb933993.aspx 
''' </remarks>
Public Class Datafiles
	Inherits Dictionary(Of Integer, Datafile)

#Region "Properties"

	Private db As Database
	Private _category As String
	Private _group As String
	''' <summary>
	''' Category under which files will be stored, set via constructor.
	''' </summary>
	''' <value></value>
	''' <returns></returns>
	''' <remarks></remarks>
	Public ReadOnly Property Category() As String
		Get
			Return _category
		End Get
	End Property
	''' <summary>
	''' Current/default group for files added to the FileStore if a group is not explicitly specified.
	''' </summary>
	''' <value></value>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Property Group() As String
		Get
			Return _group
		End Get
		Set(ByVal value As String)
			_group = value
		End Set
	End Property

#End Region

#Region "Constructor"
	''' <summary>
	''' This constructor will set the Category and Group to 'Default'
	''' </summary>
	''' <remarks></remarks>
	Public Sub New()
		_category = "Default"
		Me.Group = "Default"
	End Sub
	''' <summary>
	''' This constructor will set the Group to 'Default'
	''' </summary>
	''' <param name="category"></param>
	''' <remarks></remarks>
	Public Sub New(category As String)
		_category = category
		Me.Group = "Default"
	End Sub
	''' <summary>
	''' This constructor allows Category and Group to be set.
	''' </summary>
	''' <param name="category"></param>
	''' <param name="group"></param>
	''' <remarks></remarks>
	Public Sub New(category As String, group As String)
		_category = category
		Me.Group = group
	End Sub

#End Region

#Region "Backing store access"
	''' <summary>
	''' Loads the datafiles from the backing store.
	''' </summary>
	''' <remarks></remarks>
	Public Sub LoadDatafiles()
		Dim files As List(Of Datafile) = DatafilesBackingStore.Instance.RetrieveDataFiles(Me.Category, Me.Group)
		For Each file As Datafile In files
			Me.Add(file.ID, file)
		Next
	End Sub
	''' <summary>
	''' Save the datafiles to the backing store
	''' </summary>
	''' <remarks></remarks>
	Public Sub SaveDatafiles()
		For Each file As Datafile In Me.Values
			file.Save()
		Next
	End Sub
	''' <summary>
	''' Deletes all files in the backing store and clears the Datafiles collection.
	''' </summary>
	''' <remarks></remarks>
	Public Sub DeleteDatafiles()
		For Each file As Datafile In Me.Values
			file.Delete()
		Next
		Me.Clear()
	End Sub

#End Region

#Region "Add File Helpers"
	''' <summary>
	''' Add a file to the collection using the default group.
	''' </summary>
	''' <param name="filepath"></param>
	''' <remarks></remarks>
	Public Sub AddFile(filepath As String)
		Me.AddFile(Me.Group, filepath)
	End Sub
	''' <summary>
	''' Add a file to the collection using the given group.
	''' </summary>
	''' <param name="group"></param>
	''' <param name="filepath"></param>
	''' <remarks></remarks>
	Public Sub AddFile(group As String, filepath As String)
		'Create new datafile
		Dim newFile As New Datafile(Me.Category, group, filepath)
		'Add to collection
		Me.Add(newFile.ID, newFile)
	End Sub

#End Region

#Region "Retrieve File Helpers"
	''' <summary>
	'''  Get the Id of the file with the given parameters in the current/default Group.
	''' </summary>
	''' <param name="filename"></param>
	''' <param name="extension"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Function GetFileID(filename As String, extension As String) As Integer
		If extension.Substring(0, 1).Contains(".") Then extension = extension.Substring(1, extension.Length - 1)
		Return (From f In Me.Values Where f.Group = Me.Group And f.Filename = filename And f.Extension = extension).First.ID
	End Function
	''' <summary>
	''' Get the Id of the file with the given parameters.
	''' </summary>
	''' <param name="group"></param>
	''' <param name="filename"></param>
	''' <param name="extension"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Function GetFileID(group As String, filename As String, extension As String) As Integer
		If extension.Substring(0, 1).Contains(".") Then extension = extension.Substring(1, extension.Length - 1)
		Return (From f In Me.Values Where f.Group = group And f.Filename = filename And f.Extension = extension).First.ID
	End Function
	''' <summary>
	''' Get a file as binary.
	''' </summary>
	''' <param name="id"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Function GetBinary(id As Integer) As Byte()
		Return Me(id).Content
	End Function
	''' <summary>
	''' Get the text for a datafile.
	''' </summary>
	''' <param name="id"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Function GetText(id As Integer) As String
		Dim content() As Byte
		content = Me(id).Content
		'Return UTF8Encoding.ASCII.GetString(content) 'Note that this line leaves the BOM in the returned string
		Return Encoding.UTF8.GetString(content)	'Whereas this doesn't. Also see http://andrewmatthewthompson.blogspot.co.uk/2011/02/byte-order-mark-found-using-net.html
	End Function

#End Region

#Region "Archive functions"
	''' <summary>
	''' Write all the files to an archive with the given filename and path.
	''' </summary>
	''' <param name="filename"></param>
	''' <remarks></remarks>
	Public Sub Archive(filename As String)

        'Dim temp As String = Path.GetTempPath()
        'Using zip As ZipFile = New ZipFile
        '	For Each file As Datafile In Me.Values
        '		'Save each file to the temporary directory
        '		Dim tempfilename As String = Path.Combine(temp, file.Filename)
        '		file.SaveAs(tempfilename)
        '		'Add it to the zip
        '		zip.AddFile(tempfilename)
        '		'Delete temp
        '		IO.File.Delete(tempfilename)
        '	Next
        '	zip.Save(filename)
        'End Using

	End Sub
	''' <summary>
	''' Load this class from the given filename and path.
	''' </summary>
	''' <param name="filename"></param>
	''' <remarks></remarks>
	Public Sub Unarchive(filename As String)

	End Sub

#End Region

End Class
