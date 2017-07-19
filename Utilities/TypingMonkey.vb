Imports System.Text

''' <summary>
''' The Typing monkey generates random strings.
''' </summary>
''' <remarks>
''' See http://stackoverflow.com/questions/1546472/generate-random-strings-in-vb-net
''' </remarks>
Class TypingMonkey

    Private Const AlphanumericCharacters As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"
    Private Const NumericCharacters As String = "1234567890"
    Private Const AlphabeticCharacters As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Private random As Random

#Region "Singleton"

    Public Shared ReadOnly Instance As TypingMonkey = New TypingMonkey
    ''' <summary>
    ''' Prevent anyone from instantiating this class by making New private
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub New()
        random = New Random
    End Sub

#End Region
    ''' <summary>
    ''' Internal key bashing function.
    ''' </summary>
    Private Function TypeAway(size As Integer, characters As String) As String
        Dim builder As New StringBuilder()
        Dim ch As Char
        For i As Integer = 0 To size - 1
            ch = characters(random.[Next](0, characters.Length))
            builder.Append(ch)
        Next
        Return builder.ToString()
    End Function
    ''' <summary>
    ''' The Typing Monkey Generates a random alphabetic string with the given length.
    ''' </summary>
    ''' <param name="size">Size of the string</param>
    ''' <returns>Random string</returns>
    Public Function TypeAlphabetic(size As Integer) As String
        Return TypeAway(size, AlphabeticCharacters)
    End Function
    ''' <summary>
    ''' The Typing Monkey Generates a random alphanumeric string with the given length.
    ''' </summary>
    ''' <param name="size">Size of the string</param>
    ''' <returns>Random string</returns>
    Public Function TypeAlphanumeric(size As Integer) As String
        Return TypeAway(size, AlphaNumericCharacters)
    End Function
    ''' <summary>
    ''' The Typing Monkey Generates a random string of numbers with the given length.
    ''' </summary>
    ''' <param name="size">Size of the string</param>
    ''' <returns>Random string</returns>
    Public Function TypeNumeric(size As Integer) As String
        Return TypeAway(size, AlphaNumericCharacters)
    End Function

End Class
