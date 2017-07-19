Imports System.Configuration
Imports System.IO
Imports Crypto.Asymmetric.RSAEncryption
Imports System.Text
Imports System.Security.Cryptography

Public Enum ClientServerRole
    Server = 0
    Client = 1
End Enum

Public NotInheritable Class ClientServerEncryption

    'See https://msdn.microsoft.com/en-us/library/0cwc0x23(v=vs.85).aspx

    'The cryptographic components of the .NET Framework can be combined to create different schemes to encrypt and decrypt data.

    'A simple cryptographic scheme for encrypting and decrypting data might specify the following steps: 

    '1.Each party generates a public/private key pair.
    '2.The parties exchange their public keys.
    '3.Each party generates a secret key for TripleDES encryption, for example, and encrypts the newly created key using the other's public key.
    '4.Each party sends the data to the other and combines the other's secret key with its own, in a particular order, to create a new secret key.
    '5.The parties then initiate a conversation using symmetric encryption.

    'Creating a cryptographic scheme is not a trivial task. For more information on using cryptography, see the Cryptography topic in the Platform SDK documentation at http://msdn.microsoft.com/library.

    'See also http://dotnetslackers.com/articles/security/hashing_macs_and_digital_signatures_in_net.aspx

#Region "Client Server Encrypted communication using Containers"

    Private myrole As ClientServerRole = ClientServerRole.Server
    Private myKeyContainer As String = "MY_CONTAINER_NAME"
    Private OtherPartyContainer As String = "OTHER_PARTY_CONTAINER_NAME"
    Private mySecretKey As Byte()
    Private myCombinedKey As String
    Private PersistContainers As Boolean = False

    Private Sub New()
        'Disabled default constructor
    End Sub

    Public Sub New(keycontainername As String, otherpartycontainername As String, role As ClientServerRole)
        myrole = role
        myKeyContainer = keycontainername
        OtherPartyContainer = otherpartycontainername
    End Sub

    Public Sub New(keycontainername As String, otherpartycontainername As String, role As ClientServerRole, persist As Boolean)
        myrole = role
        myKeyContainer = keycontainername
        OtherPartyContainer = otherpartycontainername
        PersistContainers = persist
    End Sub

#Region "1.Each party creates a public/private key pair."
    ''' <summary>
    ''' Create a new asymmetric key and return the public portion.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateKey() As String
        PersistRSAKeyContainer(myKeyContainer)
        Return FetchRSAKeyXML(myKeyContainer, False)
    End Function

#End Region

#Region "2.The parties exchange their public keys."
    ''' <summary>
    ''' Store the other parties public key so we can send asymmetric encrypted messages to them.
    ''' </summary>
    ''' <param name="keyxml"></param>
    ''' <remarks></remarks>
    Public Sub StoreOtherPartyKey(keyxml As String)
        PersistRSAKeyXML(keyxml, OtherPartyContainer)
    End Sub

#End Region

#Region "3.Each party generates a secret key for encryption, and encrypts the newly created key using the other's public key."
    ''' <summary>
    ''' Generate a secret key from the given key.
    ''' </summary>
    ''' <param name="key"></param>
    ''' <remarks>Because the key and salt are predetermined it is possible to restart an encrypted conversation if the class is closed.</remarks>
    Public Sub GenerateSymmetricKey(key As String, salt As String)
        Dim saltbytes As Byte() = Encoding.ASCII.GetBytes(salt)
        Debug.Print(BitConverter.ToString(saltbytes))
        Dim myKeyGenerator As New System.Security.Cryptography.Rfc2898DeriveBytes(key, saltbytes, EncryptionSettings.Iterations)
        'mySecretKey = RSAEncrypt(myKeyGenerator.GetBytes(EncryptionSettings.KeyBitSize), OtherPartyContainer, False)
        mySecretKey = myKeyGenerator.GetBytes(EncryptionSettings.KeyBytes)
        Debug.Print("mySecretKey: " & BitConverter.ToString(mySecretKey))
    End Sub

    ''' <summary>
    ''' Generate a secret key using a randomly created key and salt.
    ''' </summary>
    ''' <remarks>These values will be recreated each time the class is instantiated so to maintain an encrypted conversation the class must be kept alive.</remarks>
    Public Sub GenerateSymmetricKey()
        Dim rng As New RNGCryptoServiceProvider()
        'Dim keybytes As Byte() = New Byte(CInt(EncryptionSettings.KeyBitSize / 8) - 1) {}
        Dim keybytes As Byte() = New Byte(EncryptionSettings.KeyBytes - 1) {}
        rng.GetBytes(keybytes)
        'Dim saltbytes As Byte() = New Byte(CInt(EncryptionSettings.BlockBitSize / 8) - 1) {}
        Dim saltbytes As Byte() = New Byte(EncryptionSettings.BlockBytes - 1) {}
        rng.GetBytes(saltbytes)
        Dim myKeyGenerator As New System.Security.Cryptography.Rfc2898DeriveBytes(keybytes, saltbytes, EncryptionSettings.Iterations)
        'mySecretKey = RSAEncrypt(myKeyGenerator.GetBytes(EncryptionSettings.KeyBitSize), OtherPartyContainer, False)
        mySecretKey = myKeyGenerator.GetBytes(EncryptionSettings.KeyBytes)
        Debug.Print("mySecretKey: " & BitConverter.ToString(mySecretKey))
    End Sub
    ''' <summary>
    ''' Encrypt my secret key using the other parties public key so it can be sent securely and decrypted at the other end.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function EncryptSymmetricKey() As String
        Dim encryptedkeybytes As Byte() = Crypto.Asymmetric.RSAEncryption.Encrypt(mySecretKey, OtherPartyContainer, True)
        Return Convert.ToBase64String(encryptedkeybytes)
    End Function

#End Region

#Region "4.Each party sends the data to the other and combines the other's secret key with its own, in a particular order, to create a new secret key."
    ''' <summary>
    ''' Decrypt the key using my private key (as the key was sent using my public key) then combine it with my key to create a combined key.
    ''' </summary>
    ''' <param name="otherpartyencryptedkey">String</param>
    ''' <remarks></remarks>
    Public Sub CreateCombinedKey(otherpartyencryptedkey As String)
        Dim otherpartyencryptedkeybytes As Byte() = Convert.FromBase64String(otherpartyencryptedkey) 'Matches the encoding used in EncryptSymmetricKey
        Dim otherpartydecryptedkeybytes As Byte() = Crypto.Asymmetric.RSAEncryption.Decrypt(otherpartyencryptedkeybytes, myKeyContainer, True)
        Debug.Print("otherpartydecryptedkeybytes: " & BitConverter.ToString(otherpartydecryptedkeybytes))
        Dim otherpartydecryptedkey As String = ASCIIEncoding.UTF8.GetString(otherpartydecryptedkeybytes)
        'We have to create the combined key differently depending on which end of the communication we are at
        Select Case myrole
            Case ClientServerRole.Server
                myCombinedKey = ASCIIEncoding.UTF8.GetString(mySecretKey) & "_" & otherpartydecryptedkey
            Case ClientServerRole.Client
                myCombinedKey = otherpartydecryptedkey & "_" & ASCIIEncoding.UTF8.GetString(mySecretKey)
        End Select
        Debug.Print("myCombinedKey: " & myCombinedKey)
    End Sub

#End Region

#Region "5.The parties then initiate a conversation using symmetric encryption."

    Public Function Encrypt(cleartext As String) As String
        Return Crypto.Symmetric.AESHMACEncryption.Encrypt(cleartext, myCombinedKey)
    End Function

    Public Function Decrypt(ciphertext As String) As String
        Return Crypto.Symmetric.AESHMACEncryption.Decrypt(ciphertext, myCombinedKey)
    End Function

#End Region

#Region "Utilities, Clean Up, etc."
    ''' <summary>
    ''' Delete RSA Key Containers.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
        If PersistContainers = False Then
            DeleteRSAKeyContainer(OtherPartyContainer)
            DeleteRSAKeyContainer(myKeyContainer)
        End If
    End Sub

#End Region

#End Region


End Class
