Imports System.Xml.Serialization

Public Class Template

#Region "Serializable properties"

    'XXXs
    <XmlArray("XXXs")> _
    <XmlArrayItem("XXX", GetType(XXX))> _
    Public XXXs As System.Collections.Generic.List(Of XXX) = New System.Collections.Generic.List(Of XXX)

    'YYY
    <XmlElement()> _
    Public YYY As String

    'ZZZ
    <XmlAttributeAttribute()> _
    Public ZZZ As Integer

#End Region

#Region "Do not serialize"
    'Properties must be tagged XMLIgnore

    'AAA
    <XmlIgnore()> _
    Public AAA As String

#End Region
End Class

Public Class XXX

#Region "Serializable properties"

    'XXX
    <XmlAttributeAttribute()> _
    Public XXX As Integer

#End Region

#Region "Do not serialize"
    'Properties must be tagged XMLIgnore

    'AAA
    <XmlIgnore()> _
    Public AAA As String

#End Region
End Class
