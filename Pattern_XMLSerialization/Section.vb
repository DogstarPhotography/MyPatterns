Imports System.Xml.Serialization
''' <summary>
''' Simple class that holds a list of name value pairs as well as data describing the list itself
''' </summary>
''' <remarks></remarks>
Public Class Section
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
    <XmlArray("NameValuePairs")> _
    <XmlArrayItem("NameValue", GetType(NameValue))> _
    Public NameValuePairs As List(Of NameValue) = New List(Of NameValue)

End Class
