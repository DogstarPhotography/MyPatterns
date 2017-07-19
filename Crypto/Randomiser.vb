Imports System.Security.Cryptography
Imports System.Text
Imports System.Threading

''' <summary>
''' Generate random strings for encryption purposes.
''' </summary>
''' <remarks>Note that none of the functions will produce the same results from the same inputs on a second run.</remarks>
Public Class Randomiser
    ''' <summary>
    ''' Generate a random string using predefined parameters.
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Shared Function Generate() As String
        Dim key As String = Shuffle("ToLongTimeString" & DateTime.Now.ToLongTimeString)
        Dim salt As String = Shuffle("ToLongDateString" & DateTime.Now.ToLongDateString)
        Dim text As String = Hash.PBKDF2(key, salt, key.Length + salt.Length)
        Return text.Substring(1, EncryptionSettings.MinimumPasswordLength)
    End Function
    ''' <summary>
    ''' Generate a random string.
    ''' </summary>
    ''' <param name="key">String</param>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Shared Function Generate(key As String) As String
        key = Shuffle(key)
        Dim salt As String = Shuffle("ToLongDateString" & DateTime.Now.ToLongDateString)
        Dim text As String = Hash.PBKDF2(key, salt, key.Length + salt.Length)
        Return text.Substring(1, EncryptionSettings.MinimumPasswordLength)
    End Function
    ''' <summary>
    ''' Generate a random string.
    ''' </summary>
    ''' <param name="key">String</param>
    ''' <param name="salt">String</param>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Shared Function Generate(key As String, salt As String) As String
        key = Shuffle(key)
        salt = Shuffle(salt)
        Dim text As String = Hash.PBKDF2(key, salt, key.Length + salt.Length)
        Return text.Substring(1, EncryptionSettings.MinimumPasswordLength)
    End Function
    ''' <summary>
    ''' Generate a random string.
    ''' </summary>
    ''' <param name="key">String</param>
    ''' <param name="salt">String</param>
    ''' <param name="length">String</param>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Shared Function Generate(key As String, salt As String, length As Integer) As String
        key = Shuffle(key)
        salt = Shuffle(salt)
        Dim lengen As Integer = length
        If lengen < key.Length + salt.Length Then lengen = key.Length + salt.Length
        Dim text As String = Hash.PBKDF2(key, salt, lengen)
        Return text.Substring(1, length)
    End Function
    ''' <summary>
    ''' Shuffle a string.
    ''' </summary>
    ''' <param name="text">String</param>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Shared Function Shuffle(text As String) As String
        Thread.Sleep(1) 'Always sleep for one millisecond to make sure that the seed for Random is not one that may have been used very recently
        Dim rnd = New Random(CInt(Date.Now.Ticks And &HFFFF))
        Dim iterations As Integer = rnd.Next(EncryptionSettings.Iterations)
        Dim interim As String = text
        For count = 1 To iterations
            interim = New String(interim.OrderBy(Function(r) rnd.[Next]).ToArray())
        Next
        Return interim
    End Function
End Class
