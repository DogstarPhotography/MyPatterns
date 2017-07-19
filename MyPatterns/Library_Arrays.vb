Public Class Library_Arrays

#Region "Convert to and from a List"
    Public Shared Function ListConvert(Of T)(ByVal Source() As T) As Generic.List(Of T)

        Dim Result As New Generic.List(Of T)

        For Index As Integer = 0 To Source.GetUpperBound(0)
            Result.Add(Source(Index))
        Next

        Return Result

    End Function

    Public Shared Function ListConvert(Of T)(ByVal Source As Generic.List(Of T)) As T()

        Dim Result(0) As T
        Dim curIndex As Integer = 0

        For Each curItem As T In Source
            ReDim Preserve Result(curIndex)
            Result(curIndex) = curItem
            curIndex = curIndex + 1
        Next

        Return Result

    End Function
#End Region

#Region "Convert to and from a Dictionary"
    Public Shared Function DictionaryConvert(Of T)(ByVal Source() As T) As Generic.Dictionary(Of Integer, T)

        Dim Result As New Generic.Dictionary(Of Integer, T)

        For Index As Integer = 0 To Source.GetUpperBound(0)
            Result.Add(Index, Source(Index))
        Next

        Return Result

    End Function

    Public Shared Function DictionaryConvert(Of T)(ByVal Source As Generic.Dictionary(Of Integer, T)) As T()

        Dim Result(Source.Count) As T

        For Each curKVP As KeyValuePair(Of Integer, T) In Source
            Result(curKVP.Key) = curKVP.Value
        Next

        Return Result

    End Function
#End Region
End Class
