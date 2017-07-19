Imports System.Xml.Serialization

Public Class Section

	Public Sub New()
		'Default constructor
	End Sub

	Public Sub New(name As String)
		Me.Name = name
	End Sub

	'Section Name
	<XmlAttributeAttribute()> Public Name As String

	'Items
	<XmlArray("Entries")> _
	<XmlArrayItem("Entry", GetType(Entry))> _
	Public Entries As System.Collections.Generic.List(Of Entry) = New System.Collections.Generic.List(Of Entry)

End Class

