''' <summary>
''' Singleton configuration data containing a dictionary of name/value pair settings
''' </summary>
''' <remarks>This use the Pattern_Singleton design pattern to provide a single shared configuration class</remarks>
Public NotInheritable Class MyConfiguration
    ''' <summary>
    ''' Shared internal variables
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared TheSingleton As MyConfiguration
    Private Shared SyncObject As New Object
    ''' <summary>
    ''' Shared public settings
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared MySettings As Dictionary(Of String, String)
    ''' <summary>
    ''' Making the Constructor private prevents instances of this class being created
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub New()
        'Need to instantiate a dictionary when this class is instantiated
        Try
            'We have to use the private variable here rather than accessing the property to prevent infinite regression...!
            MySettings = New Dictionary(Of String, String)
            'We need to load the settings at this point so that there is data
            LoadXML("DEFAULT_FILENAME.XML")
        Catch ex As Exception
            'Have to throw errors here so that the consumer is aware of the problem
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' We need to override the Finalize in order to save the data when the singleton is finished with
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub Finalize()
        Try
            MyBase.Finalize()
            'We need to save the settings at this point in case of changes
            SaveXML("DEFAULT_FILENAME.XML")
        Catch ex As Exception
            'Ignore errors
        End Try
    End Sub
    ''' <summary>
    ''' Access to settings
    ''' </summary>
    ''' <value>Dictionary(Of String, String)</value>
    ''' <returns>Dictionary(Of String, String)</returns>
    ''' <remarks>Any call to the Settings property will create an instance of the singleton if there isn't one already</remarks>
    Public Shared Property Settings() As Dictionary(Of String, String)
        Get
            Try
                'Need to prevent multiple threads from accessing this at once
                SyncLock SyncObject
                    'Create a new singleton if we don't have one
                    If TheSingleton Is Nothing Then
                        TheSingleton = New MyConfiguration
                    End If
                End SyncLock
                'Return the only instance
                Return MySettings
            Catch ex As Exception
                'Have to throw errors here so that the consumer is aware of the problem
                Throw ex
            End Try
        End Get
        Set(ByVal value As Dictionary(Of String, String))
            Try
                'Need to prevent multiple threads from accessing this at once
                SyncLock SyncObject
                    'Create a new singleton if we don't have one
                    If TheSingleton Is Nothing Then
                        TheSingleton = New MyConfiguration
                    End If
                End SyncLock
                'Use the only instance
                MySettings = value
            Catch ex As Exception
                'Have to throw errors here so that the consumer is aware of the problem
                Throw ex
            End Try
        End Set
    End Property
    ''' <summary>
    ''' Load the settings from an XML file
    ''' </summary>
    ''' <param name="XMLFilename"></param>
    ''' <remarks></remarks>
    Private Sub LoadXML(ByVal XMLFilename As String)
        'Load Items from a file
        Dim MyXMLDoc As New Xml.XmlDocument()
        Dim MyItems As Xml.XmlNodeList
        Dim curNode As Xml.XmlNode
        Dim curNodeName As String
        Dim curNodeValue As String
        Try
            'Load the file
            MyXMLDoc.PreserveWhitespace = False
            MyXMLDoc.Load(XMLFilename)
            'Get setting nodes
            MyItems = MyXMLDoc.GetElementsByTagName("Item")
            'Clear current Items
            MySettings.Clear()
            'Add a setting for each item in the nodelist
            For Each curNode In MyItems
                curNodeName = curNode.Attributes("Name").Value
                curNodeValue = curNode.Attributes("Value").Value
                MySettings.Add(curNodeName, curNodeValue)
            Next
        Catch ex As Exception
            'Have to throw errors here so that the consumer is aware of the problem
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Save the settings to an xml file
    ''' </summary>
    ''' <param name="XMLFilename"></param>
    ''' <remarks></remarks>
    Private Sub SaveXML(ByVal XMLFilename As String)
        'Save Items to a file
        Dim MyXMLDoc As New Xml.XmlDocument()
        Dim MyDeclaration As Xml.XmlDeclaration
        Dim MyItems As Xml.XmlElement
        Dim NewElement As Xml.XmlElement
        Dim NewAttribute As Xml.XmlAttribute
        Dim curItem As KeyValuePair(Of String, String)
        Try
            'Add a declaration
            MyDeclaration = MyXMLDoc.CreateXmlDeclaration("1.0", "utf-8", Nothing)
            MyXMLDoc.PrependChild(MyDeclaration)
            'Add Items element
            MyItems = MyXMLDoc.CreateElement("Items")
            MyXMLDoc.AppendChild(MyItems)
            'Add elements for each setting
            For Each curItem In MySettings
                NewElement = MyXMLDoc.CreateElement("Item")
                MyItems.AppendChild(NewElement)
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
        Catch ex As Exception
            'Have to throw errors here so that the consumer is aware of the problem
            Throw ex
        End Try
    End Sub
End Class
