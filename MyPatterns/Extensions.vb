Imports System.Data.SqlClient
Imports System.Runtime.CompilerServices
Imports System.ComponentModel

Public Module Extensions
    Private Enum SampleDescription
        <Description("Item One")> ItemOne = 1
        <Description("Item Two")> ItemTwo = 2
        <Description("Item Three has a long description")> ItemThree = 3
    End Enum
    ''' <summary>
    ''' This procedure gets the description attribute of an enum constant, if any. Otherwise it gets 
    ''' the string name of the enum member.
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks>Usage:  myenum.Member.Description()
    ''' Add the Description attribute to each member of the enumeration.</remarks>
    <Extension()> _
    Public Function Description(ByVal value As [Enum]) As String
        Dim fi As Reflection.FieldInfo = value.GetType().GetField(value.ToString())
        Dim aattr() As DescriptionAttribute = DirectCast(fi.GetCustomAttributes(GetType(DescriptionAttribute), False), DescriptionAttribute())
        If aattr.Length > 0 Then
            Return aattr(0).Description
        Else
            Return value.ToString()
        End If
    End Function
    ''' <summary>
    ''' Extension method to return a byte array from a stream.
    ''' </summary>
    ''' <param name="input">Stream</param>
    ''' <returns>Byte Array</returns>
    ''' <remarks>See http://stackoverflow.com/questions/221925/creating-a-byte-array-from-a-stream </remarks>
    <Extension> _
    Public Function ToByteArray(input As Stream) As Byte()
        Using ms As New MemoryStream()
            input.CopyTo(ms)
            Return ms.ToArray()
        End Using
    End Function
    ''' <summary>
    ''' Add an item to an array.
    ''' </summary>
    ''' <remarks>See http://stackoverflow.com/questions/18097756/fastest-way-to-add-an-item-to-an-array </remarks>
    <Extension()> _
    Public Sub Add(Of T)(ByRef arr As T(), item As T)
        Array.Resize(arr, arr.Length)
        arr(arr.Length - 1) = item
    End Sub

End Module
