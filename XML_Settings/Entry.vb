Imports System.Xml.Serialization

Public Class Entry

	Public Sub New()
		'Default constructor
	End Sub

	Public Sub New(name As String, value As String)
		Me.Name = name
		Me.Value = value
	End Sub

	<XmlAttributeAttribute()> Public Name As String
	<XmlAttributeAttribute()> Public Value As String

End Class
