Imports System.IO

Public Class INISettings
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
	''' Load sections from .INI file.
	''' </summary>
	''' <remarks></remarks>
	Public Overrides Sub Load()

		Dim myStreamReader As New StreamReader(myFilename)
		Dim curSection As Section = Nothing
		While Not myStreamReader.EndOfStream
			Dim line As String = myStreamReader.ReadLine()
			If line.Length > 0 Then	'Exclude blank lines
				If line.Contains("[") AndAlso line.Contains("]") Then
					curSection = New Section(GetINISectionName(line))
					Me.Sections.Add(curSection)
				ElseIf line.Contains("=") Then
					If curSection IsNot Nothing Then 'We must have a current section in order to have an entry. Entries before a section are invalid
						curSection.Entries.Add(New Entry(GetINIEntryName(line), GetINIEntryValue(line)))
					End If
				Else
					'Not a valid line - ignore
				End If
			End If
		End While
		myStreamReader.Close()

	End Sub
	''' <summary>
	''' Save sections to INI file.
	''' </summary>
	''' <remarks></remarks>
	Public Overrides Sub Save()

		Dim myStreamWriter As New StreamWriter(myFilename)
		For Each sect As Section In Me.Sections
			myStreamWriter.WriteLine("[" & sect.Name & "]")
			For Each ent As Entry In sect.Entries
				myStreamWriter.WriteLine(ent.Name & "=" & ent.Value)
			Next
			myStreamWriter.WriteLine("")
		Next
		myStreamWriter.Close()

	End Sub

#End Region
End Class
