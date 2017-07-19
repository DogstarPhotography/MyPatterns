Imports System.Xml.Serialization

Public Class MyNameValuePair

    '<XmlAttributeAttribute()> switches this property to be an attribute of the MyNameValuePair entry in an XML file
    '<XmlAttributeAttribute(AttributeName:="Name")> adds a named attribute 'Name'
    'This produces the following in an XML file:
    '   <MyObject Name="" Value="" />
    <XmlAttributeAttribute()> Public Name As String
    <XmlAttributeAttribute()> Public Value As String

    Public Sub New()
        'Required to allow this class to be constructed without any starting parameters
    End Sub

    Public Sub New(ByVal Name As String, ByVal Value As String)
        Me.Name = Name
        Me.Value = Value
    End Sub

End Class
