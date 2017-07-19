Imports System.IO
Imports System.Xml.Serialization
Imports System.Xml
Imports System.Text
Imports System.Runtime.Serialization.Formatters.Soap

Public Class Serialization

#Region "Generic File Serialization"
	''' <summary>
	''' Generic method to serialize an object and save it into a file
	''' </summary>
	''' <typeparam name="T"></typeparam>
	''' <param name="Filename"></param>
	''' <param name="Target"></param>
	''' <remarks></remarks>
	Public Shared Sub SerializeToFile(Of T As New)(ByVal Filename As String, ByVal Target As T)
		Dim newSerializer As New XmlSerializer(GetType(T))
		Using newWriter As New StreamWriter(Filename)
			newSerializer.Serialize(newWriter, Target)
			newWriter.Close()
		End Using
	End Sub
	''' <summary>
	''' Generic method to deserialize a file and create an object of the given type from it
	''' </summary>
	''' <typeparam name="T"></typeparam>
	''' <param name="Filename"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Function DeserializeFromFile(Of T As New)(ByVal Filename As String) As T
		Dim newT As T
		'Read file
		Dim newSerializer As New XmlSerializer(GetType(T))
		Using newReader As New StreamReader(Filename)
			newT = CType(newSerializer.Deserialize(newReader), T)
			'All done
			Return newT
		End Using
	End Function
#End Region

#Region "Generic Text Serialization"
	''' <summary>
	''' Generic method to serialize an object and return it as text
	''' </summary>
	''' <typeparam name="T"></typeparam>
	''' <param name="Target"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Function SerializeToText(Of T As New)(ByVal Target As T) As String
		Dim newSerializer As New XmlSerializer(GetType(T))
		Using newWriter As New StringWriter()
			newSerializer.Serialize(newWriter, Target)
			newWriter.Close()
			Return newWriter.ToString
		End Using
	End Function
	''' <summary>
	''' Generic method to deserialize the given text and create an object of the given type from it
	''' </summary>
	''' <typeparam name="T"></typeparam>
	''' <param name="Text"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Function DeserializeFromText(Of T As New)(ByVal Text As String) As T
		Dim newT As T
		'Read text
		Dim newSerializer As New XmlSerializer(GetType(T))
		Using newReader As New StringReader(Text)
			newT = CType(newSerializer.Deserialize(newReader), T)
			'All done
			Return newT
		End Using
	End Function
	''' <summary>
	''' Serialize an object without adding the XML declaration, etc.
	''' </summary>
	''' <param name="target"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Function SerializeElementToText(Of T As New)(target As T) As String
		Dim serializer As New XmlSerializer(GetType(T))
		'Need to serialize without namespaces to keep it clean and tidy
		Dim emptyNS As New XmlSerializerNamespaces({XmlQualifiedName.Empty})
		'Need to remove xml declaration as we will use this as part of a larger xml file
		Dim settings As New XmlWriterSettings()
		settings.OmitXmlDeclaration = True
		settings.NewLineHandling = NewLineHandling.Entitize
		settings.Indent = True
		settings.IndentChars = (ControlChars.Tab)
		Using stream As New StringWriter(), writer As XmlWriter = XmlWriter.Create(stream, settings)
			'Serialize the item to the stream using the namespace supplied
			serializer.Serialize(writer, target, emptyNS)
			'Read the stream and return it as a string
			Return stream.ToString
		End Using
	End Function
	''' <summary>
	''' Deserialize an element
	''' </summary>
	''' <typeparam name="T"></typeparam>
	''' <param name="xml"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Function DeserializeElementFromText(Of T As New)(xml As String) As T
		Return DeserializeFromText(Of T)(xml)
	End Function
#End Region

#Region "Any"

	Public Shared Function XMLToElement(xml As String) As XmlElement
		Dim doc As New XmlDocument
		doc.LoadXml(xml)
		Return doc.DocumentElement
	End Function

#End Region

#Region "Object Serialization"
	''' <summary>
	''' Serialize the given object.
	''' </summary>
	''' <param name="item"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Function SerializeObject(ByVal item As Object) As String
		'Get the type of the object and create a serializer
		Dim objecttype As Type = item.GetType
		Dim xmlSerializer As New System.Xml.Serialization.XmlSerializer(objecttype)
		Dim stringBuilder As New StringBuilder
		Dim stringWriter As New System.IO.StringWriter(stringBuilder)
		xmlSerializer.Serialize(stringWriter, item)
		Return stringBuilder.ToString()
	End Function
	''' <summary>
	''' Deserialize from xml to an object of the given type.
	''' </summary>
	''' <param name="xml"></param>
	''' <param name="objecttype"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Function DeserializeObject(ByVal xml As String, ByVal objecttype As Type) As Object
		'Extract the inner xml
		Dim xmldoc As New XmlDocument
		Dim innerxml As String
		xmldoc.LoadXml(xml)
		innerxml = xmldoc.InnerXml
		'Deserialize as an object of the given type
		Dim xmlTextReader As New System.Xml.XmlTextReader(New System.IO.StringReader(innerxml))
		Dim mySerializer = New System.Xml.Serialization.XmlSerializer(objecttype)
		Return mySerializer.Deserialize(xmlTextReader)
	End Function
	''' <summary>
	''' Serialize an object using the SOAP format.
	''' </summary>
	''' <param name="item"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Function SerializeSOAPObject(ByVal item As Object) As String
		Using ms As New MemoryStream
			Dim XMLFormatter As New SoapFormatter()
			XMLFormatter.Serialize(ms, item)
			Return Encoding.UTF8.GetString(ms.ToArray)
		End Using
	End Function
	''' <summary>
	''' Deserialize an object from SOAP format.
	''' </summary>
	''' <param name="xml"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Function DeserializeSOAPObject(ByVal xml As String) As Object
		Using ms As New MemoryStream(Encoding.UTF8.GetBytes(xml))
			Dim XMLFormatter As New SoapFormatter()
			Return XMLFormatter.Deserialize(ms)
		End Using
	End Function

#End Region

End Class
