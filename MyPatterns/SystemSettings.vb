Imports System.ComponentModel

#Region "System Settings Enumeration"
''' <summary>
''' Contains a list of all Settings used within the application as an enum. 
''' Each setting has a number of attributes applied using a SystemSettingInformation tag so that the attributes of a setting can be retrieved in code.
''' The enum as a whole is used as a key for getting and setting settings on any class that requires them - see SystemSettingClassBase for an example of procedures to add.
''' </summary>
''' <remarks>
''' Copy this file to a new application and add enum items as required. Add a GetSetting and a SetSetting to any class that needs system settings - see SystemSettingClassBase for an example of procedures to add.
''' </remarks>
Public Enum SystemSettings As Integer
	'Newly Created Settings
	<SystemSettingInformation("Support_friendly_description", "Default_Value", "Category", "Optional_Additional_Notes")> EnumItemName

	'Cut and paste the line below for new items in the enumeration
	'<SystemSettingInfo("Support_friendly_description", "Default_Value"e, "Category", "Optional_Additional_Notes")> EnumItemName
End Enum
#End Region

''' <summary>
''' Custom Attribute Class and Helper Methods for SystemSettings
''' </summary>
<System.AttributeUsage(System.AttributeTargets.Field)>
Public Class SystemSettingInformation
	Inherits System.Attribute

#Region "Properties"
	''' <summary>
	''' Explanatory description of setting
	''' </summary>
	Public Property Description As String
	''' <summary>
	''' Default Value if SetupValue is missing from table
	''' </summary>
	Public Property DefaultValue As String
	''' <summary>
	''' Category to which this setting applies
	''' </summary>
	Public Property Category As String
	''' <summary>
	''' Generic Notes field to hold additional information
	''' </summary>
	Public Property Notes As String
	''' <summary>
	''' Enum Key of this Enum
	''' </summary>
	Private Property Key As String

#End Region

#Region "Constructors"
	''' <summary>
	''' New Constructor
	''' </summary>
	''' <param name="Description">Explanatory description of Setting</param>
	''' <param name="DefaultValue">Default Value if SetupValue is missing from Table</param>
	''' <param name="Corporate">Corporate(s) to whom this setting applies</param>
	''' <param name="Notes">Optional additional information that may be useful</param>
	''' <remarks></remarks>
	Sub New(Description As String, DefaultValue As String, Corporate As String, Optional Notes As String = "")
		Me.Description = Description
		Me.DefaultValue = DefaultValue
		Me.Category = Corporate
		Me.Notes = Notes
	End Sub
	''' <summary>
	''' New Constructor
	''' </summary>
	''' <param name="Description">Explanatory description of Setting</param>
	''' <param name="DefaultValue">Default Value if SetupValue is missing from Table</param>
	''' <remarks>
	''' This constructor should only be used for Legacy Settings which have limited information at this time
	''' </remarks>
	Sub New(Description As String, DefaultValue As String)
		Me.Description = Description
		Me.DefaultValue = DefaultValue
	End Sub

#End Region

#Region "Shared Methods"
	''' <summary>
	''' Returns an Array of SystemSettingInfo objects from SystemSetting Enum
	''' </summary>    
	''' <remarks></remarks>
	Public Shared Function GetSystemSettingItems() As SystemSettingInformation()
		'Get all Enum Items
		Dim fields() As System.Reflection.FieldInfo = GetType(SystemSettings).GetFields
		Dim values As New Generic.List(Of SystemSettingInformation)
		For Each field As System.Reflection.FieldInfo In fields
			'Get Custom Attributes from Decorator
			Dim attribute() As Object = field.GetCustomAttributes(GetType(SystemSettingInformation), False)
			If attribute.Length > 0 Then
				Dim item As SystemSettingInformation = CType(attribute(0), SystemSettingInformation)
				item.Key = field.Name
				values.Add(item)
			End If
		Next
		'Return array
		Return values.ToArray
	End Function
	''' <summary>
	''' Exports all System Settings to a CSV File in the Temp Directory, 
	''' </summary>
	''' <param name="openDefaultViewer">Calls Process.Start on the file to open the default viewer</param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Function ExportToCSV(openDefaultViewer As Boolean) As String
		'Export to Temp File
		Dim exportFile As String = IO.Path.ChangeExtension(IO.Path.GetTempFileName, "csv")
		'CSV Formatted output
		Dim csvOutput As New System.Text.StringBuilder()
		csvOutput.AppendLine("Key,Description,Default Value,Category,Notes")
		Dim lineFormat As String = """{0}"",""{1}"",""{2}"",""{3}"",""{4}"",""{5}"""
		'Iterate through each setting and append to the string builder
		For Each setting As SystemSettingInformation In GetSystemSettingItems()
			'Format the CSV Line and append
			csvOutput.AppendFormat(lineFormat,
			  setting.Key,
			  setting.Description,
			  setting.DefaultValue,
			  setting.Category,
			  setting.Notes,
			  Environment.NewLine)
		Next
		'Save File and Load CSV Viewer
		IO.File.AppendAllText(exportFile, csvOutput.ToString)
		If openDefaultViewer Then Process.Start(exportFile)
		'Return Temp File Path
		Return exportFile
	End Function

#End Region

End Class

''' <summary>
''' Extension Methods for SystemSettings Enumeration
''' </summary>
''' <remarks>Uses reflection to add Custom Attributes as Properties of the SystemSetting Enum</remarks>
Public Module SystemSettingExtenders

#Region "Properties"
	''' <summary>
	''' Returns the DefaultValue Custom Attribute of this Enum
	''' </summary>
	''' <remarks>
	''' Note that this method uses Reflection, beware of performance degradation
	''' </remarks>
	<Runtime.CompilerServices.Extension()> _
	Public Function DefaultValue(ByVal value As SystemSettings) As String
		Return GetSystemSettingAttribute(value, "DefaultValue")
	End Function
	''' <summary>
	''' Returns the Corporate Custom Attribute of this Enum
	''' </summary>
	''' <remarks>
	''' Note that this method uses Reflection, beware of performance degradation
	''' </remarks>
	<Runtime.CompilerServices.Extension()> _
	Public Function Corporate(ByVal value As SystemSettings) As String
		Return GetSystemSettingAttribute(value, "Category")
	End Function
	''' <summary>
	''' Returns the Description Custom Attribute of this Enum
	''' </summary>
	''' <remarks>
	''' Note that this method uses Reflection, beware of performance degradation
	''' </remarks>
	<Runtime.CompilerServices.Extension()> _
	Public Function Description(ByVal value As SystemSettings) As String
		Return GetSystemSettingAttribute(value, "Description")
	End Function
	''' <summary>
	''' Returns the Notes Custom Attribute of this Enum
	''' </summary>
	''' <remarks>
	''' Note that this method uses Reflection, beware of performance degradation
	''' </remarks>
	<Runtime.CompilerServices.Extension()> _
	Public Function Notes(ByVal value As SystemSettings) As String
		Return GetSystemSettingAttribute(value, "Notes")
	End Function

#End Region

#Region "Private Methods"
	''' <summary>
	''' Uses Reflection to query the Custom Attributes of SystemSettingInfo and Return the results as a string
	''' </summary>
	''' <param name="attributeName"></param>
	''' <returns></returns>
	''' <remarks>
	''' </remarks>
	Private Function GetSystemSettingAttribute(ByVal value As SystemSettings, attributeName As String) As String
		Dim field As Reflection.FieldInfo = value.GetType().GetField(value.ToString())
		Dim attributeText As String = String.Empty
		'Get Custom Attributes from Decorator
		Dim attribute() As Object = field.GetCustomAttributes(GetType(SystemSettingInformation), False)
		If attribute.Length > 0 Then
			Dim item As SystemSettingInformation = CType(attribute(0), SystemSettingInformation)
			'Return the appropriate property
			Select Case attributeName.ToUpper
				Case "CATEGORY"
					attributeText = item.Category
				Case "DEFAULTVALUE"
					attributeText = item.DefaultValue
				Case "DESCRIPTION"
					attributeText = item.Description
				Case "NOTES"
					attributeText = item.Notes
			End Select
		End If
		Return attributeText
	End Function
#End Region

End Module


