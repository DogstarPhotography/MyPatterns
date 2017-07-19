Imports System.Xml
	Partial Class Catalog
		Private thisNode As XmlNode
		Public Sub New(node As XmlNode)
			thisNode = node
		End Sub
		Public ReadOnly Iterator Property Artist() As IEnumerable(Of Artist)
			Get
				For Each node As XmlNode In thisNode.SelectNodes("Artist")
					Yield New Artist(node)
				Next
			End Get
		End Property
	End Class
	Partial Class Artist
		Private thisNode As XmlNode
		Public Sub New(node As XmlNode)
			thisNode = node
		End Sub
		Public ReadOnly Iterator Property Song() As IEnumerable(Of Song)
			Get
				For Each node As XmlNode In thisNode.SelectNodes("Song")
					Yield New Song(node)
				Next
			End Get
		End Property
		Public ReadOnly Property id() As String
			Get
				Return thisNode.Attributes("id").Value
			End Get
		End Property
		Public ReadOnly Property name() As String
			Get
				Return thisNode.Attributes("name").Value
			End Get
		End Property
	End Class
	Partial Class Song
		Private thisNode As XmlNode
		Public Sub New(node As XmlNode)
			thisNode = node
		End Sub
		Public ReadOnly Property Text() As String
			Get
				Return thisNode.InnerText
			End Get
		End Property
		Public ReadOnly Property id() As String
			Get
				Return thisNode.Attributes("id").Value
			End Get
		End Property
	End Class

