Imports System.Xml.Serialization
''' <summary>
''' Simple class to represent a name value pair
''' </summary>
''' <remarks></remarks>
Public Class NameValue
    ''' <summary>
    ''' XmlAttributeAttribute specifies that this property should be serialized as an attribute
    ''' </summary>
    ''' <remarks></remarks>
    <XmlAttribute()> _
    Public Name As String
    ''' <summary>
    ''' XmlElement specifies that this property should be serialized as an ement
    ''' </summary>
    ''' <remarks></remarks>
    <XmlElement()> _
    Public Value As String
End Class
