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



'See http://stackoverflow.com/questions/129389/how-do-you-do-a-deep-copy-an-object-in-net-c-specifically/11308879#11308879

''Imports System.Collections.Generic
''Imports System.Reflection
''Imports System.ArrayExtensions

'Namespace System
'    Public NotInheritable Class ObjectExtensions
'        Private Sub New()
'        End Sub
'        Private Shared ReadOnly CloneMethod As MethodInfo = GetType([Object]).GetMethod("MemberwiseClone", BindingFlags.NonPublic Or BindingFlags.Instance)

'        <System.Runtime.CompilerServices.Extension> _
'        Public Shared Function IsPrimitive(type As Type) As Boolean
'            If type Is GetType([String]) Then
'                Return True
'            End If
'            Return (type.IsValueType And type.IsPrimitive)
'        End Function

'        <System.Runtime.CompilerServices.Extension> _
'        Public Shared Function Copy(originalObject As [Object]) As [Object]
'            Return InternalCopy(originalObject, New Dictionary(Of [Object], [Object])(New ReferenceEqualityComparer()))
'        End Function

'        Private Shared Function InternalCopy(originalObject As [Object], visited As IDictionary(Of [Object], [Object])) As [Object]
'            If originalObject Is Nothing Then
'                Return Nothing
'            End If
'            Dim typeToReflect = originalObject.[GetType]()
'            If IsPrimitive(typeToReflect) Then
'                Return originalObject
'            End If
'            If visited.ContainsKey(originalObject) Then
'                Return visited(originalObject)
'            End If
'            Dim cloneObject = CloneMethod.Invoke(originalObject, Nothing)
'            If typeToReflect.IsArray Then
'                Dim arrayType = typeToReflect.GetElementType()
'                If IsPrimitive(arrayType) = False Then
'                    Dim clonedArray As Array = DirectCast(cloneObject, Array)
'                    clonedArray.ForEach(Function(array, indices) array.SetValue(InternalCopy(clonedArray.GetValue(indices), visited), indices))

'                End If
'            End If
'            visited.Add(originalObject, cloneObject)
'            CopyFields(originalObject, visited, cloneObject, typeToReflect)
'            RecursiveCopyBaseTypePrivateFields(originalObject, visited, cloneObject, typeToReflect)
'            Return cloneObject
'        End Function

'        Private Shared Sub RecursiveCopyBaseTypePrivateFields(originalObject As Object, visited As IDictionary(Of Object, Object), cloneObject As Object, typeToReflect As Type)
'            If typeToReflect.BaseType IsNot Nothing Then
'                RecursiveCopyBaseTypePrivateFields(originalObject, visited, cloneObject, typeToReflect.BaseType)
'                CopyFields(originalObject, visited, cloneObject, typeToReflect.BaseType, BindingFlags.Instance Or BindingFlags.NonPublic, Function(info) info.IsPrivate)
'            End If
'        End Sub

'        Private Shared Sub CopyFields(originalObject As Object, visited As IDictionary(Of Object, Object), cloneObject As Object, typeToReflect As Type, Optional bindingFlags__1 As BindingFlags = BindingFlags.Instance Or BindingFlags.NonPublic Or BindingFlags.[Public] Or BindingFlags.FlattenHierarchy, Optional filter As Func(Of FieldInfo, Boolean) = Nothing)
'            For Each fieldInfo As FieldInfo In typeToReflect.GetFields(bindingFlags__1)
'                If filter IsNot Nothing AndAlso filter(fieldInfo) = False Then
'                    Continue For
'                End If
'                If IsPrimitive(fieldInfo.FieldType) Then
'                    Continue For
'                End If
'                Dim originalFieldValue = fieldInfo.GetValue(originalObject)
'                Dim clonedFieldValue = If(originalFieldValue Is Nothing, Nothing, InternalCopy(originalFieldValue, visited))
'                fieldInfo.SetValue(cloneObject, clonedFieldValue)
'            Next
'        End Sub
'        <System.Runtime.CompilerServices.Extension> _
'        Public Shared Function Copy(Of T)(original As T) As T
'            Return DirectCast(Copy(DirectCast(original, [Object])), T)
'        End Function
'    End Class

'    Public Class ReferenceEqualityComparer
'        Inherits EqualityComparer(Of [Object])
'        Public Overrides Function Equals(x As Object, y As Object) As Boolean
'            Return ReferenceEquals(x, y)
'        End Function
'        Public Overrides Function GetHashCode(obj As Object) As Integer
'            If obj Is Nothing Then
'                Return 0
'            End If
'            Return obj.GetHashCode()
'        End Function
'    End Class

'    Namespace ArrayExtensions
'        Public NotInheritable Class ArrayExtensions
'            Private Sub New()
'            End Sub
'            <System.Runtime.CompilerServices.Extension> _
'            Public Shared Sub ForEach(array As Array, action As Action(Of Array, Integer()))
'                Dim walker As New ArrayTraverse(array)
'                Do
'                    action(array, walker.Position)
'                Loop While walker.[Step]()
'            End Sub
'        End Class

'        Friend Class ArrayTraverse
'            Public Position As Integer()
'            Private maxLengths As Integer()

'            Public Sub New(array As Array)
'                maxLengths = New Integer(array.Rank - 1) {}
'                For i As Integer = 0 To array.Rank - 1
'                    maxLengths(i) = array.GetLength(i) - 1
'                Next
'                Position = New Integer(array.Rank - 1) {}
'            End Sub

'            Public Function [Step]() As Boolean
'                For i As Integer = 0 To Position.Length - 1
'                    If Position(i) < maxLengths(i) Then
'                        Position(i) += 1
'                        For j As Integer = 0 To i - 1
'                            Position(j) = 0
'                        Next
'                        Return True
'                    End If
'                Next
'                Return False
'            End Function
'        End Class
'    End Namespace






