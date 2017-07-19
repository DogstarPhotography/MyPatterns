Namespace Infrastructure
    ''' <summary>
    ''' Store Object in the session
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SessionRepository
        Implements IDictionary(Of String, Object)

        Private mySession As HttpSessionStateBase
        Private myDictionary As Dictionary(Of String, Object)

#Region "Constructors"
        ''' <summary>
        ''' Default constructor is disabled
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub New()

        End Sub

        Public Sub New(session As HttpSessionStateBase)
            mySession = session
            'Check to see if a repository already exists in the session
            If mySession.Item("SessionRepository") Is Nothing Then
                myDictionary = New Dictionary(Of String, Object)
                mySession.Add("SessionRepository", myDictionary)
            Else
                myDictionary = CType(mySession.Item("SessionRepository"), Global.System.Collections.Generic.Dictionary(Of String, Object))
            End If
        End Sub

#End Region

#Region "Implements IDictionary(Of String, Object)"

        Public Sub Add(key As String, value As Object) Implements System.Collections.Generic.IDictionary(Of String, Object).Add
            'TODO: Add any custom code here before adding to internal dictionary
            myDictionary.Add(key, value)
        End Sub

        Public Function ContainsKey(key As String) As Boolean Implements System.Collections.Generic.IDictionary(Of String, Object).ContainsKey
            Return myDictionary.ContainsKey(key)
        End Function

        Default Public Property Item(key As String) As Object Implements System.Collections.Generic.IDictionary(Of String, Object).Item
            Get
                'TODO: Add any custom code here before reading from internal dictionary
                Return myDictionary.Item(key)
            End Get
            Set(value As Object)
                'TODO: Add any custom code here before adding to internal dictionary
                myDictionary.Item(key) = value
            End Set
        End Property

        Public Function Remove(key As String) As Boolean Implements System.Collections.Generic.IDictionary(Of String, Object).Remove
            myDictionary.Remove(key)
            Return True
        End Function

        Public Function TryGetValue(key As String, ByRef value As Object) As Boolean Implements System.Collections.Generic.IDictionary(Of String, Object).TryGetValue
            Return myDictionary.TryGetValue(key, value)
        End Function

        Public ReadOnly Property Values As System.Collections.Generic.ICollection(Of Object) Implements System.Collections.Generic.IDictionary(Of String, Object).Values
            Get
                Return myDictionary.Values
            End Get
        End Property

#End Region

#Region "Implements ICollection(Of String, Object)"

        Public Sub Add(item As System.Collections.Generic.KeyValuePair(Of String, Object)) Implements System.Collections.Generic.ICollection(Of System.Collections.Generic.KeyValuePair(Of String, Object)).Add
            'TODO: Add any custom code here before adding to internal dictionary
            myDictionary.Add(item.Key, item.Value)
        End Sub

        Public Sub Clear() Implements System.Collections.Generic.ICollection(Of System.Collections.Generic.KeyValuePair(Of String, Object)).Clear
            myDictionary.Clear()
        End Sub

        Public Function Contains(item As System.Collections.Generic.KeyValuePair(Of String, Object)) As Boolean Implements System.Collections.Generic.ICollection(Of System.Collections.Generic.KeyValuePair(Of String, Object)).Contains
            Return myDictionary.ContainsKey(item.Key)
        End Function

        Public Sub CopyTo(array() As System.Collections.Generic.KeyValuePair(Of String, Object), arrayIndex As Integer) Implements System.Collections.Generic.ICollection(Of System.Collections.Generic.KeyValuePair(Of String, Object)).CopyTo
            CType(myDictionary, ICollection(Of KeyValuePair(Of String, Object))).CopyTo(array, arrayIndex)
        End Sub

        Public ReadOnly Property Count As Integer Implements System.Collections.Generic.ICollection(Of System.Collections.Generic.KeyValuePair(Of String, Object)).Count
            Get
                Return myDictionary.Count
            End Get
        End Property

        Public ReadOnly Property IsReadOnly As Boolean Implements System.Collections.Generic.ICollection(Of System.Collections.Generic.KeyValuePair(Of String, Object)).IsReadOnly
            Get
                Return False
            End Get
        End Property

        Public Function Remove(item As System.Collections.Generic.KeyValuePair(Of String, Object)) As Boolean Implements System.Collections.Generic.ICollection(Of System.Collections.Generic.KeyValuePair(Of String, Object)).Remove
            myDictionary.Remove(item.Key)
            Return True
        End Function

        Public ReadOnly Property Keys As System.Collections.Generic.ICollection(Of String) Implements System.Collections.Generic.IDictionary(Of String, Object).Keys
            Get
                Return myDictionary.Keys
            End Get
        End Property

#End Region

#Region "Implements IEnumerable(Of String, Object), IEnumerable"

        Public Function GetEnumerator() As System.Collections.Generic.IEnumerator(Of System.Collections.Generic.KeyValuePair(Of String, Object)) Implements System.Collections.Generic.IEnumerable(Of System.Collections.Generic.KeyValuePair(Of String, Object)).GetEnumerator
            Return myDictionary.GetEnumerator
        End Function

        Public Function GetEnumerator1() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
            Return myDictionary.GetEnumerator
        End Function

#End Region

    End Class

End Namespace
