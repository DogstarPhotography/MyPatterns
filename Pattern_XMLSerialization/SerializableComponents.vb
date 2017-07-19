Imports System.Xml.Serialization

Public Interface IHaveData

    Property Data1() As String
    Property Data2() As String
    Property Data3() As String

End Interface
''' <summary>
''' Contains only attributes
''' </summary>
''' <remarks>Implements IHaveData</remarks>
Public Class XXX
    Implements IHaveData

    Private myData1 As String = "TESTDATA"
    Private myData2 As String = "TESTDATA"
    Private myData3 As String = "TESTDATA"

    <XmlAttributeAttribute()> _
    Public Property Data1() As String Implements IHaveData.Data1
        Get
            Return myData1
        End Get
        Set(ByVal value As String)
            myData1 = value
        End Set
    End Property

    <XmlAttributeAttribute()> _
    Public Property Data2() As String Implements IHaveData.Data2
        Get
            Return myData2
        End Get
        Set(ByVal value As String)
            myData2 = value
        End Set
    End Property

    <XmlAttributeAttribute()> _
    Public Property Data3() As String Implements IHaveData.Data3
        Get
            Return myData3
        End Get
        Set(ByVal value As String)
            myData3 = value
        End Set
    End Property
End Class
''' <summary>
''' Mixed attributes and elements
''' </summary>
''' <remarks></remarks>
Public Class YYY

    <XmlAttributeAttribute()> _
    Public Data1 As String = "TESTDATA"
    <XmlAttributeAttribute()> _
    Public Data2 As String = "TESTDATA"
    <XmlElement()> _
    Public Data3 As String = "TESTDATA"

End Class
''' <summary>
''' Contains only elements
''' </summary>
''' <remarks></remarks>
Public Class ZZZ

    <XmlElement()> _
    Public Data1 As String = "TESTDATA"
    <XmlElement()> _
    Public Data2 As String = "TESTDATA"
    <XmlElement()> _
    Public Data3 As String = "TESTDATA"

End Class

