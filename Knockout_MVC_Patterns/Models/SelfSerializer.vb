Imports System.Xml.Serialization
Imports System.Xml
Imports System.IO
Imports System.Diagnostics.CodeAnalysis
Imports System.Runtime.Serialization

''' <summary>
''' Base class for any object that can serialize itself either normally or cleanly. i.e. without the xml declaration thus making it useful for insertion into other xml.
''' </summary>
''' <remarks>Note that self deserialization is not possible but we can implement a workaround by copying each of the properties from a source object/XML.</remarks>
<Serializable(), ExcludeFromCodeCoverage()> Public MustInherit Class SelfSerializer
    ''' <summary>
    ''' This class can serialize itself with namespace and xml declaration.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2202:Do not dispose objects multiple times")>
    Public Function Serialize() As String
        Dim serializer As New XmlSerializer(Me.GetType)
        'Need to serialize without namespaces to keep it clean and tidy
        Dim emptyNS As New XmlSerializerNamespaces({XmlQualifiedName.Empty})
        Using stream As New StringWriter(), writer As XmlWriter = XmlWriter.Create(stream)
            'Serialize the item to the stream using the namespace supplied
            serializer.Serialize(writer, Me)
            'Read the stream and return it as a string
            Return stream.ToString
        End Using
    End Function
    ''' <summary>
    ''' This class can serialize itself without namespace or xml declaration.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2202:Do not dispose objects multiple times")>
    Public Function SerializeAsElement() As String
        Dim serializer As New XmlSerializer(Me.GetType)
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
            serializer.Serialize(writer, Me, emptyNS)
            'Read the stream and return it as a string
            Return stream.ToString
        End Using
    End Function

#Region "Self deserialization"

    <XmlIgnore> Private myState As Object = Me
    ''' <summary>
    ''' Fill out the properties of this class by deserializing the given XML into an object of the same type as this class and then copying it's properties.
    ''' </summary>
    ''' <param name="xml">String</param>
    ''' <remarks>This is a shallow copy so won't read any nested objects.</remarks>
    Public Sub Deserialize(xml As String)
        Dim newMe As Object 'We don't know our own type so we have to use an object here
        'Read text
        Dim newSerializer As New XmlSerializer(Me.GetType)
        Using newReader As New StringReader(xml)
            newMe = newSerializer.Deserialize(newReader)
            myState = newMe
        End Using
        For Each propinfo In myState.GetType.GetProperties()
            Dim name = propinfo.Name
            Dim realProp = (From p In Me.GetType.GetProperties
              Where p.Name = name And p.MemberType = Reflection.MemberTypes.Property).Take(1)(0)
            'realProp.SetValue(Me, propinfo.GetValue(myState, Nothing), Nothing)
            'Replaced above line as it fails on a readonly property - need to test to ensure this doesn't have a knokc on effect elsewhere
            If realProp.CanWrite = True Then realProp.SetValue(Me, propinfo.GetValue(myState, Nothing), Nothing)
        Next
    End Sub

#End Region

End Class






