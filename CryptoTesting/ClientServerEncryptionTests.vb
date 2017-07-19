Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()> Public Class ClientServerEncryptionTests

    <TestMethod()> Public Sub CompleteClientServerEncryptionTest()

        Dim cseClient As New Crypto.ClientServerEncryption("CompleteClientServerEncryptionTest_Client", "CompleteClientServerEncryptionTest_Client_Other", Crypto.ClientServerRole.Client)
        Dim cseServer As New Crypto.ClientServerEncryption("CompleteClientServerEncryptionTest_Server", "CompleteClientServerEncryptionTest_Server_Other", Crypto.ClientServerRole.Server)

        '1.Each party generates a public/private key pair.
        Dim ClientKey As String = cseClient.CreateKey()
        Assert.IsTrue(ClientKey.Length > 0)
        Debug.Print(ClientKey)
        Dim ServerKey As String = cseServer.CreateKey()
        Assert.IsTrue(ServerKey.Length > 0)
        Debug.Print(ServerKey)

        '2.The parties exchange their public keys.
        cseClient.StoreOtherPartyKey(ServerKey)
        cseServer.StoreOtherPartyKey(ClientKey)

        '3.Each party generates a secret key and encrypts the newly created key using the other's public key.
        cseClient.GenerateSymmetricKey("Client_Key", "Client_Salt")
        cseServer.GenerateSymmetricKey("Server_Key", "Server_Salt")
        Dim EncryptedClientKey As String = cseClient.EncryptSymmetricKey()
        Assert.IsTrue(EncryptedClientKey.Length > 0)
        Debug.Print(EncryptedClientKey)
        Dim EncryptedServerKey As String = cseServer.EncryptSymmetricKey()
        Assert.IsTrue(EncryptedServerKey.Length > 0)
        Debug.Print(EncryptedServerKey)

        '4.Each party sends the data to the other and combines the other's secret key with its own, in a particular order, to create a new secret key.
        'cseClient.CreateCombinedKey(EncryptedClientKey)
        cseClient.CreateCombinedKey(EncryptedServerKey)
        cseServer.CreateCombinedKey(EncryptedClientKey)

        '5.The parties then initiate a conversation using symmetric encryption.
        Dim TestData As String = "TEST_DATA_GOES_HERE"
        Dim encrypted As String = cseClient.Encrypt(TestData)
        Dim decrypted As String = cseServer.Decrypt(encrypted)
        Assert.IsTrue(decrypted = TestData)
        TestData = "MODIFIED_TEST_DATA_GOES_HERE"
        encrypted = cseServer.Encrypt(TestData)
        decrypted = cseClient.Decrypt(encrypted)
        Assert.IsTrue(decrypted = TestData)

    End Sub

End Class