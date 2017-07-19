Imports System.Xml.Serialization
Imports System.IO
Imports System.Xml

''' <summary>
''' Example of a self serializing generic dictionary
''' </summary>
''' <typeparam name="K"></typeparam>
''' <typeparam name="V"></typeparam>
''' <remarks></remarks>
Public Class SelfSerializingDictionary(Of K As IXmlSerializable, V As IXmlSerializable)
    Inherits Dictionary(Of K, V)

#Region "Serialization"
    ''' <summary>
    ''' This class can serialize itself.
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Function Serialize() As String
        Dim keySerializer As New XmlSerializer(GetType(K))
        Dim valueSerializer As New XmlSerializer(GetType(V))
        Dim settings As New XmlWriterSettings()
        Using stream As New StringWriter(), writer As XmlWriter = XmlWriter.Create(stream)
            writer.WriteProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8""")
            writer.WriteStartElement("items")
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
            writer.WriteEndElement()
            writer.WriteEndDocument()
            writer.Flush()
            Return stream.ToString
        End Using
    End Function
    ''' <summary>
    ''' Read all the items from the given XML.
    ''' </summary>
    ''' <param name="xml">String</param>
    ''' <remarks>Clears any existing items in the dictionary before adding the items stored in the XML. Note that input XML must be valid and correct.</remarks>
    Public Sub Deserialize(xml As String)
        Dim keySerializer As New XmlSerializer(GetType(K))
        Dim valueSerializer As New XmlSerializer(GetType(V))
        Using newReader As New StringReader(xml), xmlReader As XmlReader = xmlReader.Create(newReader)
            Dim wasEmpty As Boolean = xmlReader.IsEmptyElement
            xmlReader.Read()
            If wasEmpty Then
                Return
            End If
            xmlReader.ReadStartElement("items")
            Me.Clear()
            While xmlReader.NodeType <> System.Xml.XmlNodeType.EndElement
                xmlReader.ReadStartElement("item")
                xmlReader.ReadStartElement("key")
                Dim key As K = DirectCast(keySerializer.Deserialize(xmlReader), K)
                xmlReader.ReadEndElement()
                xmlReader.ReadStartElement("value")
                Dim value As V = DirectCast(valueSerializer.Deserialize(xmlReader), V)
                xmlReader.ReadEndElement()
                Me.Add(key, value)
                xmlReader.ReadEndElement()
                xmlReader.MoveToContent()
            End While
            xmlReader.ReadEndElement()
            xmlReader.ReadEndElement()
        End Using
    End Sub

#End Region


End Class
