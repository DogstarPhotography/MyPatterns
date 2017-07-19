Imports System.Xml.Serialization
''' <summary>
''' Class that holds a list of lists
''' </summary>
''' <remarks></remarks>
<XmlRoot()> _
Public Class SectionFile
    ''' <summary>
    ''' XmlAttributeAttribute specifies that this property should be serialized as an attribute
    ''' </summary>
    ''' <remarks></remarks>
    <XmlAttributeAttribute()> _
    Public Name As String
    ''' <summary>
    ''' XmlElement specifies that this property should be serialized as an ement
    ''' </summary>
    ''' <remarks></remarks>
    <XmlElement()> _
    Public Description As String
    ''' <summary>
    ''' XmlArray enables the list to be serialized
    ''' XmlArrayItem explicitly tells the serializer to expect the given class so it can be properly writtento XML from the collection
    ''' </summary>
    ''' <remarks></remarks>
    <XmlArray("Sections")> _
    <XmlArrayItem("Section", GetType(Section))> _
    Public Sections As List(Of Section) = New List(Of Section)
    ''' <summary>
    ''' XmlIgnore specifies that this property should be ignored
    ''' </summary>
    ''' <remarks></remarks>
    <XmlIgnore()> _
    Public SecretNote As String

End Class
