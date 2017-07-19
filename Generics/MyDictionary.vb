Public Class Data
	Property Name As String
	Property Description As String
End Class
''' <summary>
''' Template for implementing IDictionary(Of TKey, TValue).
''' </summary>
''' <remarks></remarks>
Public Class MyDictionary
	Implements IDictionary(Of String, Data)

#Region "Implementation"
	'Use a generic dictionary as our internal store
	Private myDictionary As New Dictionary(Of String, Data)

#Region "Implements IDictionary(Of String, Data)"

	Public Sub Add1(key As String, value As Data) Implements System.Collections.Generic.IDictionary(Of String, Data).Add
		'TODO: Add any custom code here before adding to internal dictionary
		myDictionary.Add(key, value)
	End Sub

	Public Function ContainsKey(key As String) As Boolean Implements System.Collections.Generic.IDictionary(Of String, Data).ContainsKey
		Return myDictionary.ContainsKey(key)
	End Function

	Default Public Property Item(key As String) As Data Implements System.Collections.Generic.IDictionary(Of String, Data).Item
		Get
			'TODO: Add any custom code here before reading from internal dictionary
			Return myDictionary.Item(key)
		End Get
		Set(value As Data)
			'TODO: Add any custom code here before adding to internal dictionary
			myDictionary.Item(key) = value
		End Set
	End Property

	Public Function Remove1(key As String) As Boolean Implements System.Collections.Generic.IDictionary(Of String, Data).Remove
		myDictionary.Remove(key)
	End Function

	Public Function TryGetValue(key As String, ByRef value As Data) As Boolean Implements System.Collections.Generic.IDictionary(Of String, Data).TryGetValue
		Return myDictionary.TryGetValue(key, value)
	End Function

	Public ReadOnly Property Values As System.Collections.Generic.ICollection(Of Data) Implements System.Collections.Generic.IDictionary(Of String, Data).Values
		Get
			Return myDictionary.Values
		End Get
	End Property

#End Region

#Region "Implements ICollection(Of String, Data)"

	Public Sub Add(item As System.Collections.Generic.KeyValuePair(Of String, Data)) Implements System.Collections.Generic.ICollection(Of System.Collections.Generic.KeyValuePair(Of String, Data)).Add
		'TODO: Add any custom code here before adding to internal dictionary
		myDictionary.Add(item.Key, item.Value)
	End Sub

	Public Sub Clear() Implements System.Collections.Generic.ICollection(Of System.Collections.Generic.KeyValuePair(Of String, Data)).Clear
		myDictionary.Clear()
	End Sub

	Public Function Contains(item As System.Collections.Generic.KeyValuePair(Of String, Data)) As Boolean Implements System.Collections.Generic.ICollection(Of System.Collections.Generic.KeyValuePair(Of String, Data)).Contains
		Return myDictionary.ContainsKey(item.Key)
	End Function

	Public Sub CopyTo(array() As System.Collections.Generic.KeyValuePair(Of String, Data), arrayIndex As Integer) Implements System.Collections.Generic.ICollection(Of System.Collections.Generic.KeyValuePair(Of String, Data)).CopyTo
		CType(myDictionary, ICollection(Of KeyValuePair(Of String, Data))).CopyTo(array, arrayIndex)
	End Sub

	Public ReadOnly Property Count As Integer Implements System.Collections.Generic.ICollection(Of System.Collections.Generic.KeyValuePair(Of String, Data)).Count
		Get
			Return myDictionary.Count
		End Get
	End Property

	Public ReadOnly Property IsReadOnly As Boolean Implements System.Collections.Generic.ICollection(Of System.Collections.Generic.KeyValuePair(Of String, Data)).IsReadOnly
		Get
			Return False
		End Get
	End Property

	Public Function Remove(item As System.Collections.Generic.KeyValuePair(Of String, Data)) As Boolean Implements System.Collections.Generic.ICollection(Of System.Collections.Generic.KeyValuePair(Of String, Data)).Remove
		myDictionary.Remove(item.Key)
	End Function

	Public ReadOnly Property Keys As System.Collections.Generic.ICollection(Of String) Implements System.Collections.Generic.IDictionary(Of String, Data).Keys
		Get
			Return myDictionary.Keys
		End Get
	End Property

#End Region

#Region "Implements IEnumerable(Of String, Data), IEnumerable"

	Public Function GetEnumerator() As System.Collections.Generic.IEnumerator(Of System.Collections.Generic.KeyValuePair(Of String, Data)) Implements System.Collections.Generic.IEnumerable(Of System.Collections.Generic.KeyValuePair(Of String, Data)).GetEnumerator
		Return myDictionary.GetEnumerator
	End Function

	Public Function GetEnumerator1() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
		Return myDictionary.GetEnumerator
	End Function

#End Region

#End Region

End Class