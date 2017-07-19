Imports System.Xml.Serialization

Public Class Items

    'Items
    <XmlArray("Items")> _
    <XmlArrayItem("Item", GetType(Item))> _
    Public Items As System.Collections.Generic.List(Of Item) = New System.Collections.Generic.List(Of Item)

End Class

