Imports System.Xml
''' <summary>
''' 
''' </summary>
''' <remarks></remarks>
Public Class MyDictionary

    Public Items As Dictionary(Of String, String)

    Public Sub New()
        Items = New Dictionary(Of String, String)
    End Sub
    ''' <summary>
    ''' Load an XML file into this dictionary
    ''' </summary>
    ''' <param name="XMLFilename"></param>
    ''' <remarks></remarks>
    Public Sub LoadXML(ByVal XMLFilename As String)
        'Load Items from a file
        Dim MyXMLDoc As New Xml.XmlDocument()
        Dim MyNodes As Xml.XmlNodeList
        Dim curNode As Xml.XmlNode
        Dim curNodeName As String
        Dim curNodeValue As String
        'Load the file
        MyXMLDoc.PreserveWhitespace = False
        MyXMLDoc.Load(XMLFilename)
        'Get setting nodes
        MyNodes = MyXMLDoc.GetElementsByTagName("Item")
        'Clear current Items
        Items.Clear()
        'Add a setting for each item in the nodelist
        For Each curNode In MyNodes
            curNodeName = curNode.Attributes("Name").Value
            curNodeValue = curNode.Attributes("Value").Value
            Items.Add(curNodeName, curNodeValue)
        Next
    End Sub
    ''' <summary>
    ''' Write this dictionary into an XML file
    ''' </summary>
    ''' <param name="XMLFilename"></param>
    ''' <remarks></remarks>
    Public Sub SaveXML(ByVal XMLFilename As String)
        'Save Items to a file
        Dim MyXMLDoc As New Xml.XmlDocument()
        Dim MyDeclaration As Xml.XmlDeclaration
        Dim MyElements As Xml.XmlElement
        Dim NewElement As Xml.XmlElement
        Dim NewAttribute As Xml.XmlAttribute
        Dim curItem As KeyValuePair(Of String, String)
        'Add a declaration
        MyDeclaration = MyXMLDoc.CreateXmlDeclaration("1.0", "utf-8", Nothing)
        MyXMLDoc.PrependChild(MyDeclaration)
        'Add Items element
        MyElements = MyXMLDoc.CreateElement("Items")
        MyXMLDoc.AppendChild(MyElements)
        'Add elements for each setting
        For Each curItem In Items
            NewElement = MyXMLDoc.CreateElement("Item")
            MyElements.AppendChild(NewElement)
            'Add name attribute
            NewAttribute = MyXMLDoc.CreateAttribute("Name")
            NewAttribute.Value = curItem.Key
            NewElement.Attributes.Append(NewAttribute)
            'Add value attribute
            NewAttribute = MyXMLDoc.CreateAttribute("Value")
            NewAttribute.Value = curItem.Value
            NewElement.Attributes.Append(NewAttribute)
        Next
        'Save
        MyXMLDoc.Save(XMLFilename)
    End Sub
End Class
