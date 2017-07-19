''' <summary>
''' Inheritable class to hold name and value pair settings and allow them to be loaded or saved in XML format
''' </summary>
''' <remarks>Created by Robin G Brown, Jan 2007</remarks>
Public MustInherit Class XMLConfiguration
#Region "Configuring Settings"
    ''' <summary>
    ''' Settings store
    ''' </summary>
    ''' <remarks>This is protected to only allow access to settings to inheritors</remarks>
    Protected Settings As Dictionary(Of String, String)

    Public Sub New()
        'Create a new dictionary to hold settings
        Settings = New Dictionary(Of String, String)
    End Sub

    ''' <summary>
    ''' Load settings from an XMLDocument object
    ''' </summary>
    ''' <param name="XMLConfigDoc"></param>
    ''' <remarks></remarks>
    Public Sub LoadConfigDoc(ByVal XMLConfigDoc As Xml.XmlDocument)
        'Load settings from an XmlDocument
        Dim MyNodeList As Xml.XmlNodeList
        Dim curNode As Xml.XmlNode
        Dim curNodeName As String
        Dim curNodeValue As String
        Try
            'Get setting nodes
            MyNodeList = XMLConfigDoc.GetElementsByTagName("Setting")
            'Clear current settings
            Settings.Clear()
            'Add a setting for each item in the nodelist
            For Each curNode In MyNodeList
                curNodeName = curNode.Attributes("Name").Value
                curNodeValue = curNode.Attributes("Value").Value
                Settings.Add(curNodeName, curNodeValue)
            Next
        Catch ex As Exception
            'Throw errors
            Throw (ex)
        End Try
    End Sub

    ''' <summary>
    ''' Load settings from a file
    ''' </summary>
    ''' <param name="XMLConfigFilename"></param>
    ''' <remarks></remarks>
    Public Sub LoadConfigFile(ByVal XMLConfigFilename As String)
        'Load settings from a file
        Dim XMLConfigDoc As New Xml.XmlDocument()
        Try
            'Load the file
            XMLConfigDoc.PreserveWhitespace = False
            XMLConfigDoc.Load(XMLConfigFilename)
            'Load settings
            LoadConfigDoc(XMLConfigDoc)
        Catch ex As Exception
            'Throw errors
            Throw (ex)
        End Try
    End Sub

    ''' <summary>
    ''' Load settings from a string in XML format
    ''' </summary>
    ''' <param name="XMLConfig"></param>
    ''' <remarks></remarks>
    Public Sub LoadConfig(ByVal XMLConfig As String)
        'Load settings from a string
        Dim XMLConfigDoc As New Xml.XmlDocument()
        Try
            'Load the file
            XMLConfigDoc.PreserveWhitespace = False
            XMLConfigDoc.LoadXml(XMLConfig)
            'Load settings
            LoadConfigDoc(XMLConfigDoc)
        Catch ex As Exception
            'Throw errors
            Throw (ex)
        End Try
    End Sub

    ''' <summary>
    ''' Save settings to an XMLDocument object
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SaveConfigDoc() As Xml.XmlDocument
        'Save settings to an Xml.XmlDocument
        Dim XMLConfigDoc As New Xml.XmlDocument()
        Dim MySettingsElement As Xml.XmlElement
        Dim NewElement As Xml.XmlElement
        Dim curSetting As KeyValuePair(Of String, String)
        Try
            'Use the Library_XMLDOM functions to add XML items
            AddXmlDeclaration(XMLConfigDoc)
            'Add settings element
            MySettingsElement = AddXmlElement(XMLConfigDoc, XMLConfigDoc, "Settings")
            'Add elements for each setting
            For Each curSetting In Settings
                NewElement = AddXmlElement(XMLConfigDoc, MySettingsElement, "Setting")
                AddXmlAttribute(XMLConfigDoc, NewElement, "Name", curSetting.Key)
                AddXmlAttribute(XMLConfigDoc, NewElement, "Value", curSetting.Value)
            Next
            'Finished
            Return XMLConfigDoc
        Catch ex As Exception
            'Throw errors
            Throw (ex)
        End Try
    End Function

    ''' <summary>
    ''' Save settings to a file
    ''' </summary>
    ''' <param name="XMLConfigFilename"></param>
    ''' <remarks></remarks>
    Public Sub SaveConfigFile(ByVal XMLConfigFilename As String)
        'Save settings to a file
        Dim MyXMLDoc As Xml.XmlDocument
        Try
            'Get settings
            MyXMLDoc = SaveConfigDoc()
            'Save
            MyXMLDoc.Save(XMLConfigFilename)
        Catch ex As Exception
            'Throw errors
            Throw (ex)
        End Try
    End Sub

    ''' <summary>
    ''' Get settings as a string containing XML format data
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetConfig() As String
        'Return settings as xml string
        Dim MyXMLDoc As Xml.XmlDocument
        Dim string_writer As New IO.StringWriter()
        Dim xml_text_writer As New Xml.XmlTextWriter(string_writer)
        Try
            'Get settings
            MyXMLDoc = SaveConfigDoc()
            'Return as string
            Return XmlDocToString(MyXMLDoc)
        Catch ex As Exception
            'Throw errors
            Throw (ex)
        End Try
    End Function
#End Region
End Class
