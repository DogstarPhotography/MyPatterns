Imports System.Text
Imports System.IO
Imports System.Security.Cryptography

Namespace Symmetric
    ''' <summary>
    ''' Class that provides simple AES encryption without authentication.
    ''' </summary>
    ''' <remarks>Note that in order to use the AESEncryption class the caller must have the password and salt available. To communicate via symmetric encryption without a stored salt use the AESHMACEncryption class instead.</remarks>
    Public NotInheritable Class AESEncryption

        'See http://stackoverflow.com/questions/957388/why-are-rijndaelmanaged-and-aescryptoserviceprovider-returning-different-results/4863924#4863924 as to why AesCryptoServiceProvider is the preferred class

        ''' <summary>
        ''' Encrypt the given text using the given password and create a string that can be decrypted by AES.Decrypt.
        ''' </summary>
        ''' <param name="cleartext">String</param>
        ''' <param name="password">String</param>
        ''' <param name="salt">String</param>
        ''' <returns>String</returns>
        ''' <remarks></remarks>
        Shared Function Encrypt(cleartext As String, password As String, salt As String) As String

            'Sanity checks
            If password.Length < EncryptionSettings.MinimumPasswordLength Then Throw New Exception("Password is too short.")
            If salt.Length < EncryptionSettings.MinimumPasswordLength Then Throw New Exception("Salt is too short.")
            If cleartext.Length = 0 Then Throw New Exception("Nothing to encrypt.")

            Dim saltbytes As Byte() = Encoding.ASCII.GetBytes(salt)
            Dim gen As New System.Security.Cryptography.Rfc2898DeriveBytes(password, saltbytes)
            Dim encrypted As Byte()

            ' Create an AesCryptoServiceProvider object with the specified key and IV. 
            Using cryptoProvider As New AesCryptoServiceProvider With {
                     .KeySize = EncryptionSettings.KeyBitSize,
                     .BlockSize = EncryptionSettings.BlockBitSize,
                     .Mode = CipherMode.CBC,
                     .Padding = PaddingMode.PKCS7
                }

                Dim encryptor As ICryptoTransform
                With cryptoProvider
                    .Key = gen.GetBytes(.Key.Length)
                    .IV = gen.GetBytes(.IV.Length)
                    encryptor = .CreateEncryptor(.Key, .IV)
                End With

                ' Create the streams used for encryption. 
                Using msEncrypt As New MemoryStream()
                    Using csEncrypt As New CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write)
                        Using swEncrypt As New StreamWriter(csEncrypt)
                            'Write all data to the stream.
                            swEncrypt.Write(cleartext)
                        End Using
                        encrypted = msEncrypt.ToArray()

                    End Using
                End Using
            End Using
            Dim encryptedtext As String = Convert.ToBase64String(encrypted)
            Return encryptedtext

        End Function
        ''' <summary>
        ''' Decrypt a message encrypted using AES.Encrypt.
        ''' </summary>
        ''' <param name="encrypted">String</param>
        ''' <param name="password">String</param>
        ''' <param name="salt">String</param>
        ''' <returns>String</returns>
        ''' <remarks></remarks>
        Public Shared Function Decrypt(encrypted As String, password As String, salt As String) As String

            'Sanity checks
            If password.Length < EncryptionSettings.MinimumPasswordLength Then Throw New Exception("Password is too short.")
            If salt.Length < EncryptionSettings.MinimumPasswordLength Then Throw New Exception("Salt is too short.")

            Dim saltbytes As Byte() = Encoding.ASCII.GetBytes(salt)
            Dim gen As New System.Security.Cryptography.Rfc2898DeriveBytes(password, saltbytes)
            Dim EncryptedBytes As Byte() = Convert.FromBase64String(encrypted)

            Dim plaintext As String

            Using cryptoProvider As New AesCryptoServiceProvider With {
                     .KeySize = EncryptionSettings.KeyBitSize,
                     .BlockSize = EncryptionSettings.BlockBitSize,
                     .Mode = CipherMode.CBC,
                     .Padding = PaddingMode.PKCS7
                }

                Dim decryptor As ICryptoTransform
                With cryptoProvider
                    .Key = gen.GetBytes(.Key.Length)
                    .IV = gen.GetBytes(.IV.Length)
                    decryptor = .CreateDecryptor(.Key, .IV)
                End With

                ' Create the streams used for decryption. 
                Using msDecrypt As New MemoryStream(EncryptedBytes)

                    Using csDecrypt As New CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read)

                        Using srDecrypt As New StreamReader(csDecrypt)

                            ' Read the decrypted bytes from the decrypting stream 
                            ' and place them in a string.
                            plaintext = srDecrypt.ReadToEnd()
                        End Using
                    End Using
                End Using

            End Using

            Return plaintext

        End Function

    End Class
    ''' <summary>
    ''' Class that provides encryption using AES with HMAC for authentication.
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class AESHMACEncryption

        'See http://stackoverflow.com/questions/202011/encrypt-and-decrypt-a-string/10366194#10366194
        'and http://dotnetslackers.com/articles/security/hashing_macs_and_digital_signatures_in_net.aspx

        ''' <summary>
        ''' Encrypt the given text using the given password and create an authenticated message that can be decrypted by AESHMAC.Decrypt.
        ''' </summary>
        ''' <param name="cleartext">String</param>
        ''' <param name="password">String</param>
        ''' <returns>String</returns>
        ''' <remarks>AES encryption with HMAC.</remarks>
        Shared Function Encrypt(cleartext As String, password As String) As String

            'Sanity checks
            If password.Length < EncryptionSettings.MinimumPasswordLength Then Throw New Exception("Password is too short.")
            If cleartext.Length = 0 Then Throw New Exception("Nothing to encrypt.")

            'Generate two keys and salts
            Dim pairs As List(Of KeySaltPair) = KeyGeneration.GenerateKeySaltPair(password, 2)
            Dim kspCrypt As KeySaltPair = pairs.Item(0)
            Dim kspAuth As KeySaltPair = pairs.Item(1)
            'Debug.Print("kspCrypt.Key" & ASCIIEncoding.UTF8.GetString(kspCrypt.Key) & vbCrLf)
            'Debug.Print("kspAuth.Key" & ASCIIEncoding.UTF8.GetString(kspAuth.Key) & vbCrLf)

            Dim encrypted As Byte()
            Dim signed As Byte()

            ' Create an AesCryptoServiceProvider object with the specified key and IV. 
            Using cryptoProvider As New AesCryptoServiceProvider With {
                     .KeySize = EncryptionSettings.KeyBitSize,
                     .BlockSize = EncryptionSettings.BlockBitSize,
                     .Mode = CipherMode.CBC,
                     .Padding = PaddingMode.PKCS7
                }

                'Use random IV
                cryptoProvider.GenerateIV()
                Dim iv As Byte() = cryptoProvider.IV

                'Create encryptor
                Using encryptor As ICryptoTransform = cryptoProvider.CreateEncryptor(kspCrypt.Key, iv)
                    ' Create the streams used for encryption. 
                    Using msEncrypt As New MemoryStream()
                        Using csEncrypt As New CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write)
                            Using swEncrypt As New StreamWriter(csEncrypt)
                                'Write all data to the stream.
                                swEncrypt.Write(cleartext)
                            End Using
                            encrypted = msEncrypt.ToArray()

                        End Using
                    End Using
                End Using

                'Assemble encrypted message and add authentication
                Using hmac As New HMACSHA256(kspAuth.Key)
                    Using encryptedStream = New MemoryStream()
                        Using binaryWriter = New BinaryWriter(encryptedStream)
                            'Prepend Salts
                            binaryWriter.Write(kspCrypt.Salt)
                            binaryWriter.Write(kspAuth.Salt)
                            'Prepend IV
                            binaryWriter.Write(iv)
                            'Write Ciphertext
                            binaryWriter.Write(encrypted)
                            binaryWriter.Flush()
                            'Authenticate all data
                            Dim tag = hmac.ComputeHash(encryptedStream.ToArray())
                            'Postpend tag
                            binaryWriter.Write(tag)
                        End Using
                        signed = encryptedStream.ToArray()
                    End Using
                End Using
            End Using

            'Debug.Print("Message: " & ASCIIEncoding.UTF8.GetString(signed) & vbCrLf)

            Dim encryptedtext As String = Convert.ToBase64String(signed)
            Return encryptedtext

        End Function
        ''' <summary>
        ''' Decrypt a message encrypted using AESHMAC.Encrypt.
        ''' </summary>
        ''' <param name="encrypted">String</param>
        ''' <param name="password">String</param>
        ''' <returns>String</returns>
        ''' <remarks>AES encryption with HMAC.</remarks>
        Public Shared Function Decrypt(encrypted As String, password As String) As String

            'Sanity checks
            If password.Length < EncryptionSettings.MinimumPasswordLength Then Throw New Exception("Password is too short.")

            Dim encryptedMessage = Convert.FromBase64String(encrypted)
            'Debug.Print("Message: " & ASCIIEncoding.UTF8.GetString(encryptedMessage) & vbCrLf)

            Dim plaintext As String

            Dim kspCrypt As New KeySaltPair
            kspCrypt.Salt = New Byte(EncryptionSettings.SaltBitSize \ 8 - 1) {}
            Dim kspAuth As New KeySaltPair
            kspAuth.Salt = New Byte(EncryptionSettings.SaltBitSize \ 8 - 1) {}

            Dim SaltsLength As Integer = kspCrypt.Salt.Length + kspAuth.Salt.Length

            'Grab Salts from message
            Array.Copy(encryptedMessage, 0, kspCrypt.Salt, 0, kspCrypt.Salt.Length)
            Array.Copy(encryptedMessage, kspCrypt.Salt.Length, kspAuth.Salt, 0, kspAuth.Salt.Length)

            kspCrypt.Key = KeyGeneration.GenerateKey(password, kspCrypt.Salt)
            kspAuth.Key = KeyGeneration.GenerateKey(password, kspAuth.Salt)
            'Debug.Print("kspCrypt.Key" & ASCIIEncoding.UTF8.GetString(kspCrypt.Key) & vbCrLf)
            'Debug.Print("kspAuth.Key" & ASCIIEncoding.UTF8.GetString(kspAuth.Key) & vbCrLf)

            Using hmac As New HMACSHA256(kspAuth.Key)
                Dim sentTag = New Byte(hmac.HashSize \ 8 - 1) {}
                'Calculate Tag
                Dim calcTag = hmac.ComputeHash(encryptedMessage, 0, encryptedMessage.Length - sentTag.Length)
                Dim ivLength As Integer = (EncryptionSettings.BlockBitSize \ 8)

                'if message length is too small just return null
                If encryptedMessage.Length < sentTag.Length + SaltsLength + ivLength Then Throw New Exception("Message is too short to be a valid message.")

                'Grab Sent Tag
                Array.Copy(encryptedMessage, encryptedMessage.Length - sentTag.Length, sentTag, 0, sentTag.Length)

                'Compare Tag with constant time comparison
                Dim compare As Integer = 0
                For index As Integer = 0 To sentTag.Length - 1
                    compare = compare Or sentTag(index) Xor calcTag(index)
                Next

                'if message doesn't authenticate throw an error
                If compare <> 0 Then Throw New Exception("Message does not pass authentication.")

                Using cryptoProvider As New AesCryptoServiceProvider() With {
                     .KeySize = EncryptionSettings.KeyBitSize,
                     .BlockSize = EncryptionSettings.BlockBitSize,
                     .Mode = CipherMode.CBC,
                     .Padding = PaddingMode.PKCS7
                }

                    'Grab IV from message
                    Dim iv = New Byte(ivLength - 1) {}
                    Array.Copy(encryptedMessage, SaltsLength, iv, 0, iv.Length)

                    Using decrypter As ICryptoTransform = cryptoProvider.CreateDecryptor(kspCrypt.Key, iv)
                        Using plainTextStream As New MemoryStream()
                            Using decrypterStream As New CryptoStream(plainTextStream, decrypter, CryptoStreamMode.Write)
                                Using binaryWriter As New BinaryWriter(decrypterStream)
                                    'Decrypt Cipher Text from Message
                                    binaryWriter.Write(encryptedMessage, SaltsLength + iv.Length, encryptedMessage.Length - SaltsLength - iv.Length - sentTag.Length)
                                End Using
                            End Using
                            'Return Plain Text
                            plaintext = Encoding.UTF8.GetString(plainTextStream.ToArray())
                        End Using
                    End Using
                End Using
            End Using

            Return plaintext

        End Function

    End Class

End Namespace