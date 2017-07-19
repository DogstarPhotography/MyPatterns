''' <summary>
''' A class that builds a CSV string
''' </summary>
''' <remarks></remarks>
Public Class CSVBuilder

    Private MyStrings As Generic.SortedList(Of Integer, String)

    Public Sub New()
        MyStrings = New Generic.SortedList(Of Integer, String)
    End Sub
    ''' <summary>
    ''' Append the given string to the end of the CSV list
    ''' </summary>
    ''' <param name="Item"></param>
    ''' <remarks>Will replace any commas in Item with spaces to prevent bad data problems</remarks>
    Public Sub Append(ByVal Item As String)
        Try
            Dim Clean As String = ""
            'Replace commas
            Clean = Item.Replace(",", " ")
            MyStrings.Add(MyStrings.Count, Clean)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Return the comma separated values that have previously been appended
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Value() As String
        Get
            Try
                Dim Result As String = ""
                For Each curKVP As KeyValuePair(Of Integer, String) In MyStrings
                    Result = Result & curKVP.Value & ","
                Next
                'Trim last comma
                If Result.Length > 0 Then Result = Result.Substring(0, (Result.Length - 1))
                Return Result
            Catch ex As Exception
                Throw ex
            End Try
        End Get
    End Property

End Class

