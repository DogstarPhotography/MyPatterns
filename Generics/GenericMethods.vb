Public Class GenericMethods
    ''' <summary>
    ''' Convert a list of one type to an array of another type
    ''' </summary>
    ''' <typeparam name="L"></typeparam>
    ''' <typeparam name="A"></typeparam>
    ''' <param name="Items"></param>
    ''' <param name="Conversion"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ConvertListToArray(Of L, A)(ByVal Items As Generic.List(Of L), ByVal Conversion As Converter(Of L, A)) As A()

        'Use converter delegate to convert list into another list
        Dim ConvertedList As Generic.List(Of A) = Items.ConvertAll(Conversion)
        'Create an array to hold the data
        Dim ReturnArray(ConvertedList.Count - 1) As A
        'Copy the list to the array
        ConvertedList.CopyTo(ReturnArray)
        'All done
        Return ReturnArray

    End Function
    ''' <summary>
    ''' Convert an array of one type to a list of another type
    ''' </summary>
    ''' <typeparam name="A"></typeparam>
    ''' <typeparam name="L"></typeparam>
    ''' <param name="Items"></param>
    ''' <param name="Conversion"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ConvertArrayToList(Of A, L)(ByVal Items() As A, ByVal Conversion As Converter(Of A, L)) As Generic.List(Of L)

        'Use converter delegate to convert array into another array
        Dim ConvertedArray() As L = Array.ConvertAll(Items, Conversion)
        'Create new list from array (we can do this because the array implements IEnumerable)
        Dim ReturnList As New Generic.List(Of L)(ConvertedArray)
        'All done
        Return ReturnList

    End Function
End Class
