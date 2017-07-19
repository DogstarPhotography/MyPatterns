Imports System.Xml

Public Class XMLGenerator

#Region "Singleton"

    Public Shared ReadOnly Instance As XMLGenerator = New XMLGenerator
    ''' <summary>
    ''' Prevent anyone from instantiating this class by making New private
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub New()

    End Sub

#End Region

    Public Function GetRandomXMLElement() As XmlElement
        Dim doc As New XmlDocument
        Dim element As XmlElement = doc.CreateElement(TypingMonkey.Instance.TypeAlphabetic(6))
        Return element
    End Function

    Public Function GetRandomXMLStructure(depth As Integer) As XmlElement
        'Fix bounds
        If depth < 1 Then depth = 1
        If depth > 5 Then depth = 5
        Dim doc As New XmlDocument
        Dim root As XmlElement = doc.CreateElement(TypingMonkey.Instance.TypeAlphabetic(6))
        Dim elements(0 To 4) As XmlElement
        For level As Integer = 0 To depth - 1
            elements(level) = Me.GetRandomXMLElement()
        Next
        For level As Integer = depth - 1 To 1 Step -1
            elements(level - 1).AppendChild(elements(level))
        Next
        root.AppendChild(elements(0))
        Return root
    End Function

End Class
