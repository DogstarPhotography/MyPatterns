Imports System.IO
Imports System.Xml.Xsl
Imports System.Xml
Imports System.Xml.XPath

Public Class XSLTransformer

    Public Function Transform(xsltext As String, xmltext As String) As String

        'Read xsl and jump forward to stylesheet section
        Dim reader As XmlReader = XmlReader.Create(New StringReader(xsltext))
        reader.ReadToDescendant("xsl:stylesheet")

        'Load xml and create navigator
        Dim xmldoc As New XmlDocument
        xmldoc.LoadXml(xmltext)
        Dim navigator As XPathNavigator = xmldoc.CreateNavigator

        'Create transform
        Dim transformer As New XslCompiledTransform()
        transformer.Load(reader)

        'Apply xsl to xml to create xfdf - We have to muck about with streams here in order to get the byte array we actually want
        Dim xmlout As String
        Dim xfdf As Byte() = Nothing
        Using stream As New MemoryStream
            transformer.Transform(navigator, Nothing, stream)
            'Store byte array version of data
            xfdf = stream.ToArray
            Dim sr As New StreamReader(stream)
            xmlout = sr.ReadToEnd
        End Using

        Return xmlout

    End Function

End Class
