Option Strict Off 'Required because C# coders can't write anything with it on...
Imports System.Xml.Serialization
Imports System.IO
Imports System.Xml
Imports System.Runtime.Serialization

''' <summary>
''' Serializable Dictionary based on this post: http://weblogs.asp.net/pwelter34/444961
''' </summary>
''' <typeparam name="K"></typeparam>
''' <typeparam name="V"></typeparam>
''' <remarks></remarks>
<Serializable>
Public Class SerializableDictionary(Of K As IXmlSerializable, V As IXmlSerializable)
    Inherits Dictionary(Of K, V)
    Implements IXmlSerializable

#Region "Constructors"
    Public Sub New()
        MyBase.New()
    End Sub

    'Public Sub New(dictionary As IDictionary)
    '    MyBase.New(dictionary)
    'End Sub

    'Public Sub New(comparer As IEqualityComparer)
    '    MyBase.New(comparer)
    'End Sub

    Public Sub New(capacity As Integer)
        MyBase.New(capacity)
    End Sub

    Public Sub New(dictionary As IDictionary, comparer As IEqualityComparer)
        MyBase.New(dictionary, comparer)
    End Sub

    Public Sub New(capacity As Integer, comparer As IEqualityComparer)
        MyBase.New(capacity, comparer)
    End Sub

    Protected Sub New(info As SerializationInfo, context As StreamingContext)
        MyBase.New(info, context)
    End Sub
#End Region

#Region "IXmlSerializable"

    Public Function GetSchema() As Schema.XmlSchema Implements IXmlSerializable.GetSchema
        Return Nothing
    End Function

    Public Sub ReadXml(reader As XmlReader) Implements IXmlSerializable.ReadXml
        Dim keySerializer As New XmlSerializer(GetType(K))
        Dim valueSerializer As New XmlSerializer(GetType(V))
        Dim wasEmpty As Boolean = reader.IsEmptyElement
        reader.Read()
        If wasEmpty Then
            Return
        End If
        While reader.NodeType <> System.Xml.XmlNodeType.EndElement
            reader.ReadStartElement("item")
            reader.ReadStartElement("key")
            Dim key As K = DirectCast(keySerializer.Deserialize(reader), K)
            reader.ReadEndElement()
            reader.ReadStartElement("value")
            Dim value As V = DirectCast(valueSerializer.Deserialize(reader), V)
            reader.ReadEndElement()
            Me.Add(key, value)
            reader.ReadEndElement()
            reader.MoveToContent()
        End While
        reader.ReadEndElement()
    End Sub

    Public Sub WriteXml(writer As XmlWriter) Implements IXmlSerializable.WriteXml
        Dim keySerializer As New XmlSerializer(GetType(K))
        Dim valueSerializer As New XmlSerializer(GetType(V))
        For Each key As K In Me.Keys
            writer.WriteStartElement("item")
            writer.WriteStartElement("key")
            keySerializer.Serialize(writer, key)
            writer.WriteEndElement()
            writer.WriteStartElement("value")
            Dim value As V = Me(key)
            valueSerializer.Serialize(writer, value)
            writer.WriteEndElement()
            writer.WriteEndElement()
            writer.Flush()
        Next
        writer.Flush()
    End Sub

#End Region

End Class
