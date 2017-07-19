Imports System.Xml.Serialization

'---------------------------------------------------------------------------------------------
'   Cut, paste, and replace
'---------------------------------------------------------------------------------------------
''XXXs
'<XmlArray("XXXs")> _
'<XmlArrayItem("XXX", GetType(XXX))> _
'Public XXXs As System.Collections.Generic.List(Of XXX) = New System.Collections.Generic.List(Of XXX)
'---------------------------------------------------------------------------------------------
''YYY
'<XmlElement()> _
'Public YYY As Integer
'---------------------------------------------------------------------------------------------
''ZZZ
'<XmlAttributeAttribute()> _
'Public ZZZ As Integer
'---------------------------------------------------------------------------------------------
''' <summary>
''' This class contains several array elements showing different ways of creating the same XML
''' </summary>
''' <remarks></remarks>
Public Class SerializableComplexClass

    'XXXs using XMLArray and a list of XXX
    <XmlArray("XXXsArray")> _
    <XmlArrayItem("XXX", GetType(XXX))> _
    Public XXXsArray As System.Collections.Generic.List(Of XXX) = New System.Collections.Generic.List(Of XXX)

    'XXXs using XMLElement and the XXXs class below that contains a list of XXX
    <XmlElement()> _
    Public XXXsElement As XXXs = New XXXs

    'YYYs using XMLArray and a list of YYY
    <XmlArray("YYYsArray")> _
    <XmlArrayItem("YYY", GetType(YYY))> _
    Public YYYsArray As System.Collections.Generic.List(Of YYY) = New System.Collections.Generic.List(Of YYY)

    'YYYs using XMLElement and the YYYs class below that contains a list of YYY
    <XmlElement()> _
    Public YYYsElement As YYYs = New YYYs

    'ZZZs using XMLArray and a list of ZZZ
    <XmlArray("ZZZsArray")> _
    <XmlArrayItem("ZZZ", GetType(ZZZ))> _
    Public ZZZsArray As System.Collections.Generic.List(Of ZZZ) = New System.Collections.Generic.List(Of ZZZ)

    'ZZZs using XMLElement and the ZZZs class below that contains a list of ZZZ
    <XmlElement()> _
    Public ZZZsElement As ZZZs = New ZZZs

End Class

Public Class XXXs

    'XXX - Using the XmlElement markup renders the list as a flat sequence.
    <XmlElement()> _
    Public XXX As System.Collections.Generic.List(Of XXX) = New System.Collections.Generic.List(Of XXX)

End Class

Public Class YYYs

    'YYY
    <XmlElement()> _
    Public YYY As System.Collections.Generic.List(Of YYY) = New System.Collections.Generic.List(Of YYY)

End Class

Public Class ZZZs

    'ZZZ
    <XmlElement()> _
    Public ZZZ As System.Collections.Generic.List(Of ZZZ) = New System.Collections.Generic.List(Of ZZZ)

End Class