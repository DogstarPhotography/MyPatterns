Option Strict Off

Imports System.Runtime.Serialization
Imports System.IO
Imports System.Xml

Public Module SerializationExtensions
    ''' <summary>
    ''' Serialize a dictionary or pretty much any object.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="obj"></param>
    ''' <returns></returns>
    ''' <remarks>See http://stackoverflow.com/questions/1124597/why-isnt-there-an-xml-serializable-dictionary-in-net </remarks>
    <System.Runtime.CompilerServices.Extension>
    Public Function Serialize(Of T)(obj As T) As String
        Dim serializer = New DataContractSerializer(obj.[GetType]())
        Using writer = New StringWriter()
            Using stm = XmlWriter.Create(writer) 'New XmlTextWriter(writer)
                serializer.WriteObject(stm, obj)
                Return writer.ToString()
            End Using
        End Using
    End Function
    ''' <summary>
    ''' Deserialize a dictionary or pretty much any object.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="serialized"></param>
    ''' <returns></returns>
    ''' <remarks>See http://stackoverflow.com/questions/1124597/why-isnt-there-an-xml-serializable-dictionary-in-net</remarks>
    <System.Runtime.CompilerServices.Extension>
    Public Function Deserialize(Of T)(serialized As String) As T
        Dim serializer = New DataContractSerializer(GetType(T))
        Using reader = New StringReader(serialized)
            Using stm = XmlReader.Create(reader) ' New XmlTextReader(reader)
                Return DirectCast(serializer.ReadObject(stm), T)
            End Using
        End Using
    End Function

End Module
