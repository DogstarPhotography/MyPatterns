Imports System.Security.Cryptography
Imports System.Text

Namespace Symmetric
    ''' <summary>
    ''' Simple class to store key and salt byte arrays together.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class KeySaltPair

        Public Key As Byte()
        Public Salt As Byte()

        ''' <summary>
        ''' Default constructor.
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            'Do nothing
        End Sub
        ''' <summary>
        ''' Constructor that creates key and salt byte arrays from the given key.
        ''' </summary>
        ''' <param name="key"></param>
        ''' <remarks></remarks>
        Public Sub New(key As String)
            Using generator As New Rfc2898DeriveBytes(key, EncryptionSettings.SaltBitSize \ 8, EncryptionSettings.Iterations)
                Me.Salt = generator.Salt
                Me.Key = generator.GetBytes(EncryptionSettings.KeyBitSize \ 8)
            End Using
        End Sub

    End Class
    ''' <summary>
    ''' Class that provides key generation for symmetric encryption.
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class KeyGeneration
        ''' <summary>
        ''' Generate a key byte array of the given length from the given key string and salt.
        ''' </summary>
        ''' <param name="key">String</param>
        ''' <param name="salt">String</param>
        ''' <param name="length">Integer</param>
        ''' <returns>Byte Array</returns>
        ''' <remarks></remarks>
        Public Shared Function GenerateKey(key As String, salt As String, length As Integer) As Byte()
            'Sanity checks
            If key.Length < EncryptionSettings.MinimumPasswordLength Then Throw New Exception("Key is too short.")
            If salt.Length < EncryptionSettings.MinimumPasswordLength Then Throw New Exception("Salt is too short.")
            'Generate
            Dim saltbytes As Byte() = Encoding.ASCII.GetBytes(salt)
            Dim gen As New Rfc2898DeriveBytes(key, saltbytes, EncryptionSettings.Iterations)
            Dim generated As Byte() = gen.GetBytes(length)
            Debug.Print("GenerateKey: " & BitConverter.ToString(generated))
            Return generated
        End Function
        ''' <summary>
        ''' Generate a key byte array from the given key string and salt.
        ''' </summary>
        ''' <param name="key">String</param>
        ''' <param name="salt">Byte Array</param>
        ''' <returns>Byte Array</returns>
        ''' <remarks></remarks>
        Public Shared Function GenerateKey(key As String, salt As Byte()) As Byte()
            'Sanity checks
            If key.Length < EncryptionSettings.MinimumPasswordLength Then Throw New Exception("Key is too short.")
            If salt.Length < EncryptionSettings.MinimumPasswordLength \ 8 Then Throw New Exception("Salt is too short.")
            'Generate
            Dim gen As New Rfc2898DeriveBytes(key, salt, EncryptionSettings.Iterations)
            Dim generated As Byte() = gen.GetBytes(EncryptionSettings.KeyBytes)
            Debug.Print("GenerateKey: " & BitConverter.ToString(generated))
            Return generated
        End Function
        ''' <summary>
        ''' Generate a key and a random salt from the given key string.
        ''' </summary>
        ''' <param name="key">String</param>
        ''' <returns>KeySaltPair</returns>
        ''' <remarks></remarks>
        Public Shared Function GenerateKeySaltPair(key As String) As KeySaltPair
            'Sanity checks
            If key.Length < EncryptionSettings.MinimumPasswordLength Then Throw New Exception("Key is too short.")
            'Generate
            Dim ksp As New KeySaltPair
            Using generator As New Rfc2898DeriveBytes(key, EncryptionSettings.SaltBitSize \ 8, EncryptionSettings.Iterations) 'Note the \ in order to return an integer result as opposed to a / which returns a double
                ksp.Salt = generator.Salt
                Debug.Print("KSP.Salt: " & BitConverter.ToString(ksp.Salt))
                ksp.Key = generator.GetBytes(EncryptionSettings.KeyBytes)
                Debug.Print("KSP.Key: " & BitConverter.ToString(ksp.Key))
            End Using
            Return ksp
        End Function
        ''' <summary>
        ''' Generate multiple key and random salt pairs from the given key string.
        ''' </summary>
        ''' <param name="key">String</param>
        ''' <returns>KeySaltPair</returns>
        ''' <remarks></remarks>
        Public Shared Function GenerateKeySaltPair(key As String, count As Integer) As List(Of KeySaltPair)
            'Sanity checks
            If key.Length < EncryptionSettings.MinimumPasswordLength Then Throw New Exception("Key is too short.")
            'Work out how many bytes we need
            Dim totalbytes As Integer = EncryptionSettings.KeyBytes * count
            Dim keybytes As Byte()
            Dim saltbytes As Byte()
            'Generate
            Using generator As New Rfc2898DeriveBytes(key, EncryptionSettings.SaltBytes, EncryptionSettings.Iterations)
                saltbytes = generator.Salt
                keybytes = generator.GetBytes(totalbytes)
                Debug.Print("Salt: " & BitConverter.ToString(saltbytes))
                Debug.Print("Key: " & BitConverter.ToString(keybytes))
            End Using
            Dim pairs As New List(Of KeySaltPair)
            For item = 0 To count - 1
                Dim chunk(EncryptionSettings.KeyBytes - 1) As Byte
                For index = 0 To EncryptionSettings.KeyBytes - 1
                    chunk(index) = keybytes((item * EncryptionSettings.KeyBytes) + index)
                Next
                Debug.Print("Chunk " & item & ": " & BitConverter.ToString(chunk))
                Dim ksp As New KeySaltPair
                ksp.Salt = saltbytes
                ksp.Key = chunk
                pairs.Add(ksp)
            Next
            Return pairs
        End Function

    End Class

End Namespace