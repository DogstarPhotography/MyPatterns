'Option Strict Off

''Source: http://stackoverflow.com/questions/202011/encrypt-and-decrypt-a-string/10366194#10366194

'' * This work (Modern Encryption of a String C#, by James Tuley), 
'' * identified by James Tuley, is free of known copyright restrictions.
'' * https://gist.github.com/4336842
'' * http://creativecommons.org/publicdomain/mark/1.0/ 

'Imports System.IO
'Imports System.Security.Cryptography
'Imports System.Text

'Namespace Encryption
'    Public NotInheritable Class AESThenHMAC
'        Private Sub New()
'        End Sub
'        Private Shared ReadOnly Random As RandomNumberGenerator = RandomNumberGenerator.Create()

'        'Preconfigured Encryption Parameters
'        Public Shared ReadOnly BlockBitSize As Integer = 128
'        Public Shared ReadOnly KeyBitSize As Integer = 256

'        'Preconfigured Password Key Derivation Parameters
'        Public Shared ReadOnly SaltBitSize As Integer = 64
'        Public Shared ReadOnly Iterations As Integer = 10000
'        Public Shared ReadOnly MinPasswordLength As Integer = 12

'        ''' <summary>
'        ''' Helper that generates a random key on each call.
'        ''' </summary>
'        ''' <returns></returns>
'        Public Shared Function NewKey() As Byte()
'            Dim key = New Byte(KeyBitSize / 8 - 1) {}
'            Random.GetBytes(key)
'            Return key
'        End Function

'        ''' <summary>
'        ''' Simple Encryption (AES) then Authentication (HMAC) for a UTF8 Message.
'        ''' </summary>
'        ''' <param name="secretMessage">The secret message.</param>
'        ''' <param name="cryptKey">The crypt key.</param>
'        ''' <param name="authKey">The auth key.</param>
'        ''' <param name="nonSecretPayload">(Optional) Non-Secret Payload.</param>
'        ''' <returns>
'        ''' Encrypted Message
'        ''' </returns>
'        ''' <exception cref="System.ArgumentException">Secret Message Required!;secretMessage</exception>
'        ''' <remarks>
'        ''' Adds overhead of (Optional-Payload + BlockSize(16) + Message-Padded-To-Blocksize +  HMac-Tag(32)) * 1.33 Base64
'        ''' </remarks>
'        Public Shared Function SimpleEncrypt(secretMessage As String, cryptKey As Byte(), authKey As Byte(), Optional nonSecretPayload As Byte() = Nothing) As String
'            If String.IsNullOrEmpty(secretMessage) Then
'                Throw New ArgumentException("Secret Message Required!", "secretMessage")
'            End If

'            Dim plainText = Encoding.UTF8.GetBytes(secretMessage)
'            Dim cipherText = SimpleEncrypt(plainText, cryptKey, authKey, nonSecretPayload)
'            Return Convert.ToBase64String(cipherText)
'        End Function

'        ''' <summary>
'        ''' Simple Authentication (HMAC) then Decryption (AES) for a secrets UTF8 Message.
'        ''' </summary>
'        ''' <param name="encryptedMessage">The encrypted message.</param>
'        ''' <param name="cryptKey">The crypt key.</param>
'        ''' <param name="authKey">The auth key.</param>
'        ''' <param name="nonSecretPayloadLength">Length of the non secret payload.</param>
'        ''' <returns>
'        ''' Decrypted Message
'        ''' </returns>
'        ''' <exception cref="System.ArgumentException">Encrypted Message Required!;encryptedMessage</exception>
'        Public Shared Function SimpleDecrypt(encryptedMessage As String, cryptKey As Byte(), authKey As Byte(), Optional nonSecretPayloadLength As Integer = 0) As String
'            If String.IsNullOrWhiteSpace(encryptedMessage) Then
'                Throw New ArgumentException("Encrypted Message Required!", "encryptedMessage")
'            End If

'            Dim cipherText = Convert.FromBase64String(encryptedMessage)
'            Dim plainText = SimpleDecrypt(cipherText, cryptKey, authKey, nonSecretPayloadLength)
'            Return If(plainText Is Nothing, Nothing, Encoding.UTF8.GetString(plainText))
'        End Function

'        ''' <summary>
'        ''' Simple Encryption (AES) then Authentication (HMAC) of a UTF8 message
'        ''' using Keys derived from a Password (PBKDF2).
'        ''' </summary>
'        ''' <param name="secretMessage">The secret message.</param>
'        ''' <param name="password">The password.</param>
'        ''' <param name="nonSecretPayload">The non secret payload.</param>
'        ''' <returns>
'        ''' Encrypted Message
'        ''' </returns>
'        ''' <exception cref="System.ArgumentException">password</exception>
'        ''' <remarks>
'        ''' Significantly less secure than using random binary keys.
'        ''' Adds additional non secret payload for key generation parameters.
'        ''' </remarks>
'        Public Shared Function SimpleEncryptWithPassword(secretMessage As String, password As String, Optional nonSecretPayload As Byte() = Nothing) As String
'            If String.IsNullOrEmpty(secretMessage) Then
'                Throw New ArgumentException("Secret Message Required!", "secretMessage")
'            End If

'            Dim plainText = Encoding.UTF8.GetBytes(secretMessage)
'            Dim cipherText = SimpleEncryptWithPassword(plainText, password, nonSecretPayload)
'            Return Convert.ToBase64String(cipherText)
'        End Function

'        ''' <summary>
'        ''' Simple Authentication (HMAC) and then Descryption (AES) of a UTF8 Message
'        ''' using keys derived from a password (PBKDF2). 
'        ''' </summary>
'        ''' <param name="encryptedMessage">The encrypted message.</param>
'        ''' <param name="password">The password.</param>
'        ''' <param name="nonSecretPayloadLength">Length of the non secret payload.</param>
'        ''' <returns>
'        ''' Decrypted Message
'        ''' </returns>
'        ''' <exception cref="System.ArgumentException">Encrypted Message Required!;encryptedMessage</exception>
'        ''' <remarks>
'        ''' Significantly less secure than using random binary keys.
'        ''' </remarks>
'        Public Shared Function SimpleDecryptWithPassword(encryptedMessage As String, password As String, Optional nonSecretPayloadLength As Integer = 0) As String
'            If String.IsNullOrWhiteSpace(encryptedMessage) Then
'                Throw New ArgumentException("Encrypted Message Required!", "encryptedMessage")
'            End If

'            Dim cipherText = Convert.FromBase64String(encryptedMessage)
'            Dim plainText = SimpleDecryptWithPassword(cipherText, password, nonSecretPayloadLength)
'            Return If(plainText Is Nothing, Nothing, Encoding.UTF8.GetString(plainText))
'        End Function

'        Public Shared Function SimpleEncrypt(secretMessage As Byte(), cryptKey As Byte(), authKey As Byte(), Optional nonSecretPayload As Byte() = Nothing) As Byte()
'            'User Error Checks
'            If cryptKey Is Nothing OrElse cryptKey.Length <> KeyBitSize / 8 Then
'                Throw New ArgumentException([String].Format("Key needs to be {0} bit!", KeyBitSize), "cryptKey")
'            End If

'            If authKey Is Nothing OrElse authKey.Length <> KeyBitSize / 8 Then
'                Throw New ArgumentException([String].Format("Key needs to be {0} bit!", KeyBitSize), "authKey")
'            End If

'            If secretMessage Is Nothing OrElse secretMessage.Length < 1 Then
'                Throw New ArgumentException("Secret Message Required!", "secretMessage")
'            End If

'            'non-secret payload optional
'            nonSecretPayload = If(nonSecretPayload, New Byte() {})

'            Dim cipherText As Byte()
'            Dim iv As Byte()

'            Using aes As New AesManaged() With { _
'                 .KeySize = KeyBitSize, _
'                 .BlockSize = BlockBitSize, _
'                 .Mode = CipherMode.CBC, _
'                 .Padding = PaddingMode.PKCS7 _
'            }

'                'Use random IV
'                aes.GenerateIV()
'                iv = aes.IV

'                Using encrypter As ICryptoTransform = aes.CreateEncryptor(cryptKey, iv)
'                    Using cipherStream As New MemoryStream()
'                        Using cryptoStream As New CryptoStream(cipherStream, encrypter, CryptoStreamMode.Write)
'                            Using binaryWriter As New BinaryWriter(cryptoStream)
'                                'Encrypt Data
'                                binaryWriter.Write(secretMessage)
'                            End Using
'                        End Using

'                        cipherText = cipherStream.ToArray()
'                    End Using

'                End Using
'            End Using

'            'Assemble encrypted message and add authentication
'            Using hmac As New HMACSHA256(authKey)
'                Using encryptedStream = New MemoryStream()
'                    Using binaryWriter = New BinaryWriter(encryptedStream)
'                        'Prepend non-secret payload if any
'                        binaryWriter.Write(nonSecretPayload)
'                        'Prepend IV
'                        binaryWriter.Write(iv)
'                        'Write Ciphertext
'                        binaryWriter.Write(cipherText)
'                        binaryWriter.Flush()

'                        'Authenticate all data
'                        Dim tag = hmac.ComputeHash(encryptedStream.ToArray())
'                        'Postpend tag
'                        binaryWriter.Write(tag)
'                    End Using
'                    Return encryptedStream.ToArray()
'                End Using
'            End Using

'        End Function

'        Public Shared Function SimpleDecrypt(encryptedMessage As Byte(), cryptKey As Byte(), authKey As Byte(), Optional nonSecretPayloadLength As Integer = 0) As Byte()

'            'Basic Usage Error Checks
'            If cryptKey Is Nothing OrElse cryptKey.Length <> KeyBitSize / 8 Then
'                Throw New ArgumentException([String].Format("CryptKey needs to be {0} bit!", KeyBitSize), "cryptKey")
'            End If

'            If authKey Is Nothing OrElse authKey.Length <> KeyBitSize / 8 Then
'                Throw New ArgumentException([String].Format("AuthKey needs to be {0} bit!", KeyBitSize), "authKey")
'            End If

'            If encryptedMessage Is Nothing OrElse encryptedMessage.Length = 0 Then
'                Throw New ArgumentException("Encrypted Message Required!", "encryptedMessage")
'            End If

'            Using hmac As New HMACSHA256(authKey)
'                Dim sentTag = New Byte(hmac.HashSize / 8 - 1) {}
'                'Calculate Tag
'                Dim calcTag = hmac.ComputeHash(encryptedMessage, 0, encryptedMessage.Length - sentTag.Length)
'                Dim ivLength = (BlockBitSize / 8)

'                'if message length is to small just return null
'                If encryptedMessage.Length < sentTag.Length + nonSecretPayloadLength + ivLength Then
'                    Return Nothing
'                End If

'                'Grab Sent Tag
'                Array.Copy(encryptedMessage, encryptedMessage.Length - sentTag.Length, sentTag, 0, sentTag.Length)

'                'Compare Tag with constant time comparison
'                Dim compare = 0
'                For i As Integer = 0 To sentTag.Length - 1
'                    compare = compare Or sentTag(i) Xor calcTag(i)
'                Next

'                'if message doesn't authenticate return null
'                If compare <> 0 Then
'                    Return Nothing
'                End If

'                Using aes As New AesManaged() With { _
'                     .KeySize = KeyBitSize, _
'                     .BlockSize = BlockBitSize, _
'                     .Mode = CipherMode.CBC, _
'                     .Padding = PaddingMode.PKCS7 _
'                }

'                    'Grab IV from message
'                    Dim iv = New Byte(ivLength - 1) {}
'                    Array.Copy(encryptedMessage, nonSecretPayloadLength, iv, 0, iv.Length)

'                    Using decrypter As ICryptoTransform = aes.CreateDecryptor(cryptKey, iv)
'                        Using plainTextStream As New MemoryStream()
'                            Using decrypterStream As New CryptoStream(plainTextStream, decrypter, CryptoStreamMode.Write)
'                                Using binaryWriter As New BinaryWriter(decrypterStream)
'                                    'Decrypt Cipher Text from Message
'                                    binaryWriter.Write(encryptedMessage, nonSecretPayloadLength + iv.Length, encryptedMessage.Length - nonSecretPayloadLength - iv.Length - sentTag.Length)
'                                End Using
'                            End Using
'                            'Return Plain Text
'                            Return plainTextStream.ToArray()
'                        End Using
'                    End Using
'                End Using
'            End Using
'        End Function

'        Public Shared Function SimpleEncryptWithPassword(secretMessage As Byte(), password As String, Optional nonSecretPayload As Byte() = Nothing) As Byte()
'            nonSecretPayload = If(nonSecretPayload, New Byte() {})

'            'User Error Checks
'            If String.IsNullOrWhiteSpace(password) OrElse password.Length < MinPasswordLength Then
'                Throw New ArgumentException([String].Format("Must have a password of at least {0} characters!", MinPasswordLength), "password")
'            End If

'            If secretMessage Is Nothing OrElse secretMessage.Length = 0 Then
'                Throw New ArgumentException("Secret Message Required!", "secretMessage")
'            End If

'            Dim payload As Byte() = New Byte(((SaltBitSize / 8) * 2) + (nonSecretPayload.Length - 1)) {}

'            Array.Copy(nonSecretPayload, payload, nonSecretPayload.Length)
'            Dim payloadIndex As Integer = nonSecretPayload.Length

'            Dim cryptKey As Byte()
'            Dim authKey As Byte()
'            'Use Random Salt to prevent pre-generated weak password attacks.
'            Using generator As New Rfc2898DeriveBytes(password, SaltBitSize / 8, Iterations)
'                Dim salt = generator.Salt

'                'Generate Keys
'                cryptKey = generator.GetBytes(KeyBitSize / 8)

'                'Create Non Secret Payload
'                Array.Copy(salt, 0, payload, payloadIndex, salt.Length)
'                payloadIndex += salt.Length
'            End Using

'            'Deriving separate key, might be less efficient than using HKDF, 
'            'but now compatible with RNEncryptor which had a very similar wireformat and requires less code than HKDF.
'            Using generator As New Rfc2898DeriveBytes(password, SaltBitSize / 8, Iterations)
'                Dim salt = generator.Salt

'                'Generate Keys
'                authKey = generator.GetBytes(KeyBitSize / 8)

'                'Create Rest of Non Secret Payload
'                Array.Copy(salt, 0, payload, payloadIndex, salt.Length)
'            End Using

'            Return SimpleEncrypt(secretMessage, cryptKey, authKey, payload)
'        End Function

'        Public Shared Function SimpleDecryptWithPassword(encryptedMessage As Byte(), password As String, Optional nonSecretPayloadLength As Integer = 0) As Byte()
'            'User Error Checks
'            If String.IsNullOrWhiteSpace(password) OrElse password.Length < MinPasswordLength Then
'                Throw New ArgumentException([String].Format("Must have a password of at least {0} characters!", MinPasswordLength), "password")
'            End If

'            If encryptedMessage Is Nothing OrElse encryptedMessage.Length = 0 Then
'                Throw New ArgumentException("Encrypted Message Required!", "encryptedMessage")
'            End If

'            Dim cryptSalt = New Byte(SaltBitSize / 8 - 1) {}
'            Dim authSalt = New Byte(SaltBitSize / 8 - 1) {}

'            'Grab Salt from Non-Secret Payload
'            Array.Copy(encryptedMessage, nonSecretPayloadLength, cryptSalt, 0, cryptSalt.Length)
'            Array.Copy(encryptedMessage, nonSecretPayloadLength + cryptSalt.Length, authSalt, 0, authSalt.Length)

'            Dim cryptKey As Byte()
'            Dim authKey As Byte()

'            'Generate crypt key
'            Using generator As New Rfc2898DeriveBytes(password, cryptSalt, Iterations)
'                cryptKey = generator.GetBytes(KeyBitSize / 8)
'            End Using
'            'Generate auth key
'            Using generator As New Rfc2898DeriveBytes(password, authSalt, Iterations)
'                authKey = generator.GetBytes(KeyBitSize / 8)
'            End Using

'            Return SimpleDecrypt(encryptedMessage, cryptKey, authKey, cryptSalt.Length + authSalt.Length + nonSecretPayloadLength)
'        End Function
'    End Class
'End Namespace


