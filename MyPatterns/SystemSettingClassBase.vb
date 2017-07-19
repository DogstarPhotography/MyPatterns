''' <summary>
''' Adds get and set system settings procedures to a class.
''' </summary>
''' <remarks></remarks>
Public MustInherit Class SystemSettingClassBase
	''' <summary>
	''' Internal store.
	''' </summary>
	''' <remarks></remarks>
	Private mySystemSettings As New Dictionary(Of SystemSettings, String)
	''' <summary>
	''' Get the value of a system setting.
	''' </summary>
	''' <param name="setting"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Overridable Function GetSystemSetting(setting As SystemSettings) As String
		If mySystemSettings.ContainsKey(setting) = False Then
			mySystemSettings.Add(setting, setting.DefaultValue)	'Use the default value from the SystemSettings SystemSettingInformation attribute
		End If
		Return mySystemSettings.Item(setting)
	End Function
	''' <summary>
	''' Set the value of a system setting.
	''' </summary>
	''' <param name="setting"></param>
	''' <param name="value"></param>
	''' <remarks></remarks>
	Public Overridable Sub SetSystemSetting(setting As SystemSettings, value As String)
		If mySystemSettings.ContainsKey(setting) = False Then
			mySystemSettings.Add(setting, value)
		Else
			mySystemSettings.Item(setting) = value
		End If
	End Sub
	''' <summary>
	''' Any class that uses system settings must also have a way to persist them.
	''' </summary>
	''' <remarks></remarks>
	Public MustOverride Sub StoreSystemSettings()

End Class
