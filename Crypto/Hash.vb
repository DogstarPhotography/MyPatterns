Imports System.Text
Imports System.Security.Cryptography

Public Class Hash
    ''' <summary>
    ''' Generate a hash of the desired length in characters from the given text and salt.
    ''' </summary>
    ''' <param name="text">String</param>
    ''' <param name="salt">String</param>
    ''' <param name="length">Integer</param>
    ''' <returns>String</returns>
    ''' <remarks>The length must be at least equal to the size of the text plus the salt.</remarks>
    Public Shared Function PBKDF2(text As String, salt As String, length As Integer) As String
        'Sanity checks
        If text.Length < EncryptionSettings.MinimumPasswordLength Then Throw New Exception("Key is too short.")
        If salt.Length < EncryptionSettings.MinimumPasswordLength Then Throw New Exception("Salt is too short.")
        If length < (text.Length + salt.Length) Then Throw New Exception("Length is too short")
        'Generate
        Dim saltbytes As Byte() = Encoding.ASCII.GetBytes(salt)
        Dim gen As New Rfc2898DeriveBytes(text, saltbytes, EncryptionSettings.Iterations)
        Dim generated As Byte() = gen.GetBytes(length)
        Debug.Print("GenerateKey: " & BitConverter.ToString(generated))
        'Return a string of the requested size
        Return Convert.ToBase64String(generated).Substring(1, length)
    End Function
    ''' <summary>
    ''' Hash the plain text using SHA256
    ''' </summary>
    ''' <param name="plaintext">String</param>
    ''' <returns>String</returns>
    ''' <remarks>Returned string will always be 44 characters long.</remarks>
    Public Shared Function SHA256(plaintext As String) As String
        Dim plainbytes As Byte() = Encoding.ASCII.GetBytes(plaintext)
        Using csp As New SHA256CryptoServiceProvider
            Dim hashedbytes As Byte() = csp.ComputeHash(plainbytes)
            Return Convert.ToBase64String(hashedbytes)
        End Using
    End Function
    ''' <summary>
    ''' Hash the plain text using SHA512
    ''' </summary>
    ''' <param name="plaintext">String</param>
    ''' <returns>String</returns>
    ''' <remarks>Returned string will always be 88 characters long.</remarks>
    Public Shared Function SHA512(plaintext As String) As String
        Dim plainbytes As Byte() = Encoding.ASCII.GetBytes(plaintext)
        Using csp As New SHA512CryptoServiceProvider
            Dim hashedbytes As Byte() = csp.ComputeHash(plainbytes)
            Return Convert.ToBase64String(hashedbytes)
        End Using
    End Function

End Class
