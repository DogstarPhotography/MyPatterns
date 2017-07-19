Imports System.Xml.Serialization

#Region "XMLAnyElement and XMLText"
''' <summary>
''' Demonstrates the use of XMLAnyElement and XMLText attributes.
''' </summary>
''' <remarks>
''' This class allows us to serialize/deserialize XML with mixed content. See the XMLAnyElementSample.xml for XML that can be used.
''' </remarks>
<Serializable()> Public Class Content
    ''' <summary>
    ''' Default constructor.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        MyBase.New()
        Me.anyField = New List(Of System.Xml.XmlElement)()
    End Sub

#Region "Elements"

    Private anyField As List(Of System.Xml.XmlElement)
    ''' <summary>
    ''' Content of this page that is XMLElements.
    ''' </summary>
    ''' <value>List of XmlElement.</value>
    ''' <returns>List of XmlElement.</returns>
    ''' <remarks>See http://stackoverflow.com/questions/18061069/how-to-serialize-a-sequence-of-any-elements-without-a-containing-element </remarks>
    <XmlAnyElementAttribute()>
    Public Property Any() As List(Of System.Xml.XmlElement)
        Get
            Return Me.anyField
        End Get
        Set(value As List(Of System.Xml.XmlElement))
            Me.anyField = value
        End Set
    End Property

    Private myText As String = ""
    ''' <summary>
    ''' Content of this element that is text.
    ''' </summary>
    ''' <value>String</value>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    <XmlTextAttribute>
    Public Property Text() As String
        Get
            Return myText
        End Get
        Set(ByVal value As String)
            myText = value
        End Set
    End Property

#End Region

End Class

#End Region

#Region "Serializing lists"
''' <summary>
''' Demonstrates the different ways to decorate a list.
''' </summary>
''' <remarks>See http://msdn.microsoft.com/en-us/library/2baksw0z(v=vs.90).aspx </remarks>
Public Class ListSerializationDemo
    Inherits SelfSerializer
    ''' <summary>
    ''' Leaving off any attributes renders the list...
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Undecoratedlist As List(Of NameValue)
    ''' <summary>
    ''' Using the XMLElement attribute renders the list as a flat set of elements with no container.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <XmlElement> Public Property Flatlist As List(Of NameValue)

    Public Property ContainedList As List(Of NameValue)

End Class

#End Region

#Region "Serializing empty lists using Specified"
''' <summary>
''' This class demonstrates the use of a [property]Specified property to force the XMLSerializer to avoid rendering an empty list property.
''' </summary>
''' <remarks></remarks>
Public Class SpecifiedDemo
    Inherits SelfSerializer
    ''' <summary>
    ''' We can create a New list in either the constructor or explicitly which means that we do not need to deal with lazy loading issues.
    ''' </summary>
    ''' <remarks></remarks>
    Private myData As New List(Of NameValue)
    ''' <summary>
    ''' The [property]Specified property must match the name of the list property and return True if the list is to be serialized and False if it is not.
    ''' </summary>
    ''' <value>Boolean</value>
    ''' <returns>Boolean</returns>
    ''' <remarks></remarks>
    <XmlIgnore> Public Property DataSpecified() As Boolean
        Get
            If myData.Count > 0 Then Return True
            Return False
        End Get
        Set(ByVal value As Boolean)
            'Ignore
        End Set
    End Property
    ''' <summary>
    ''' Simple serializable list property.
    ''' </summary>
    ''' <value>List(Of NameValue)</value>
    ''' <returns>List(Of NameValue)</returns>
    ''' <remarks>There is no need to add lazy loading code as this list will be initialized in the constructor.</remarks>
    <XmlElement> Public Property Data() As List(Of NameValue)
        Get
            Return myData
        End Get
        Set(ByVal value As List(Of NameValue))
            myData = value
        End Set
    End Property

End Class

#End Region