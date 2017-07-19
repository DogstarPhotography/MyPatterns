''' <summary>
''' Base/abstract class for the various types of settings stores. All stores use a Section > Entry > Value object model.
''' The various Create...Settings functions are a factory to produce the different versions based on the backing store required.
''' </summary>
''' <remarks></remarks>
Public MustInherit Class Settings

#Region "Abstract Base and Inheritance"

	''' <summary>
	''' Load the settings from a backing store.
	''' </summary>
	''' <remarks></remarks>
	Public MustOverride Sub Load()
	''' <summary>
	''' Save the settings to a backing store
	''' </summary>
	''' <remarks></remarks>
	Public MustOverride Sub Save()
	''' <summary>
	''' Filename/connection string/details about the backing store.
	''' </summary>
	''' <value></value>
	''' <returns></returns>
	''' <remarks></remarks>
	Protected Friend MustOverride Property BackingStore As String

#Region "Basic operations"
	Private mySettings As New List(Of Section)
	''' <summary>
	''' Object model access: Section > Entry > Value
	''' </summary>
	''' <value></value>
	''' <returns></returns>
	''' <remarks></remarks>
	Protected Friend Property Sections As System.Collections.Generic.List(Of Section)
		Get
			Return mySettings
		End Get
		Set(value As System.Collections.Generic.List(Of Section))
			mySettings = value
		End Set
	End Property
    ''' <summary>
    ''' Return the given section
    ''' </summary>
    ''' <param name="name">String</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function GetSection(name As String) As Section
        Dim sect = (From s In Me.Sections Where s.Name = name).FirstOrDefault
        If sect Is Nothing Then
            Me.Sections.Add(New Section(name))
        End If
        Return sect
    End Function
    ''' <summary>
    ''' Add an entry with a value to the given section.
    ''' </summary>
    ''' <param name="section"></param>
    ''' <param name="entry"></param>
    ''' <param name="value"></param>
    ''' <remarks></remarks>
	Public Overridable Sub Add(section As String, entry As String, value As String)
		Dim sect = (From s In Me.Sections Where s.Name = section).FirstOrDefault
		If sect IsNot Nothing Then
			Dim ent = (From e In sect.Entries Where e.Name = entry).FirstOrDefault
			If ent IsNot Nothing Then
				'Update entry
				ent.Value = value
			Else
				'Add new entry
				sect.Entries.Add(New Entry(entry, value))
			End If
		Else
			'Add new section and entry
			Dim newSection As New Section(section)
			newSection.Entries.Add(New Entry(entry, value))
			Me.Sections.Add(newSection)
		End If
	End Sub
	''' <summary>
	''' Fetch the value for an entry in the given section. If the entry or section does not exist it is created using the default value supplied.
	''' </summary>
	''' <param name="section"></param>
	''' <param name="entry"></param>
	''' <param name="default"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Overridable Function Fetch(section As String, entry As String, Optional [default] As String = "") As String
		Dim sect = (From s In Me.Sections Where s.Name = section).FirstOrDefault
		If sect IsNot Nothing Then
			Dim ent = (From e In sect.Entries Where e.Name = entry).FirstOrDefault
			If ent IsNot Nothing Then
				Return ent.Value
			Else
				'Add new entry
				Dim newEntry As New Entry(entry, [default])
				sect.Entries.Add(newEntry)
				Return [default]
			End If
		Else
			'Add new section and entry
			Dim newSection As New Section(section)
			Dim newEntry As New Entry(entry, [default])
			newSection.Entries.Add(newEntry)
			Me.Sections.Add(newSection)
			Return [default]
		End If
	End Function
	''' <summary>
	''' Removes the entry from the given section. Does not remove a section under any circumstances.
	''' </summary>
	''' <param name="section"></param>
	''' <param name="entry"></param>
	''' <remarks></remarks>
	Public Overridable Sub Remove(section As String, entry As String)
		Dim sect = (From s In Me.Sections Where s.Name = section).FirstOrDefault
		If sect IsNot Nothing Then
			Dim ent = (From e In sect.Entries Where e.Name = entry).FirstOrDefault
			If ent IsNot Nothing Then
				sect.Entries.Remove(ent)
			End If
		End If

	End Sub

#End Region
#End Region

#Region "Factory"

	Public Shared Function CreateINISettings(ByVal filename As String) As INISettings
		Dim instance As New INISettings(filename)
		instance.Load()
		Return instance
	End Function

	Public Shared Function CreateXMLSettings(ByVal filename As String) As XMLSettings
		Dim instance As New XMLSettings(filename)
		instance.Load()
		Return instance
	End Function

	Public Shared Function ConvertINIToXML(ByVal INIFilename As String) As XMLSettings
		Dim ini As New INISettings(INIFilename)
		ini.Load()
		Dim xml As New XMLSettings()
		xml.Sections = ini.Sections
		Return xml
	End Function

	Public Shared Sub ConvertINIToXML(ByVal INIFilename As String, ByVal XMLFilename As String)
		Dim ini As New INISettings(INIFilename)
		ini.Load()
		Dim xml As New XMLSettings()
		xml.Sections = ini.Sections
		xml.BackingStore = XMLFilename
		xml.Save()
	End Sub

#End Region

#Region "INI Helper functions"
	''' <summary>
	''' Return the XXX from [XXX].
	''' </summary>
	''' <param name="text"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Protected Shared Function GetINISectionName(text As String) As String
		Dim result As String = text.Trim
		result = result.Replace("[", "")
		result = result.Replace("]", "")
		Return result.Trim
	End Function
	''' <summary>
	''' Return the XXX from XXX=YYY.
	''' </summary>
	''' <param name="text"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Protected Shared Function GetINIEntryName(text As String) As String
		Dim result As String = text.Trim
		Dim idx As Integer = result.IndexOf("=")
		Return result.Substring(0, idx)
	End Function
	''' <summary>
	''' Return the YYY from XXX=YYY.
	''' </summary>
	''' <param name="text"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Protected Shared Function GetINIEntryValue(text As String) As String
		Dim result As String = text.Trim
		Dim idx As Integer = result.IndexOf("=")
		Return result.Substring(idx + 1, (result.Length - idx - 1))
	End Function

#End Region

End Class
