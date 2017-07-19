Imports System.Xml.Serialization
Imports System.IO
Imports System.Xml

Public Class XMLSettings
	Inherits Settings

#Region "MustOverride Properties"

	Private myFilename As String

	Protected Friend Overrides Property BackingStore As String
		Get
			Return myFilename
		End Get
		Set(value As String)
			myFilename = value
		End Set
	End Property

#End Region

#Region "Backing store specific overrides"
	''' <summary>
	''' Default constructor so that we can create new empty settings
	''' </summary>
	''' <remarks></remarks>
	Protected Friend Sub New()
	End Sub
	''' <summary>
	''' Constructor requires filename.
	''' </summary>
	''' <param name="filename"></param>
	''' <remarks></remarks>
	Protected Friend Sub New(filename As String)
		myFilename = filename
	End Sub
	''' <summary>
	''' Load sections from .XML file.
	''' </summary>
	''' <remarks></remarks>
	Public Overrides Sub Load()
		Try
			Dim xmlRoot As New XmlRootAttribute()
			xmlRoot.ElementName = "Sections"
			'Read file
			Dim newSerializer As New XmlSerializer(GetType(List(Of Section)), xmlRoot)
			Dim newReader As New StreamReader(myFilename)
			Me.Sections = CType(newSerializer.Deserialize(newReader), List(Of Section))
		Catch invopex As InvalidOperationException ', xmlex As XmlException
			'Bad xml - return empty sections
			Me.Sections = New List(Of Section)
		Catch ex As Exception
			Throw ex 'Pass all other errors up
		End Try
	End Sub
	''' <summary>
	''' Save sections to .XML file.
	''' </summary>
	''' <remarks></remarks>
	Public Overrides Sub Save()

		Dim xmlRoot As New XmlRootAttribute()
		xmlRoot.ElementName = "Sections"
		'xmlRoot.Namespace = "http://www.cpandl.com"
		'xmlRoot.IsNullable = True

		Dim newWriter As New StreamWriter(myFilename)
		Dim newSerializer As New XmlSerializer(GetType(List(Of Section)), xmlRoot)
		newSerializer.Serialize(newWriter, Me.Sections)
		newWriter.Close()

	End Sub

#End Region

End Class
