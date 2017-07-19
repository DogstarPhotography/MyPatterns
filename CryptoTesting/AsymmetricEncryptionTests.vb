Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports System.IO
Imports Crypto.Asymmetric
Imports System.Configuration
Imports SSC = System.Security.Cryptography

<TestClass()> Public Class AsymmetricEncryptionTests

    <TestMethod()> Public Sub EncryptDecryptTest()
        Dim cleartext As String = "TEST_DATA_GOES_HERE"
        Dim filename As String = ConfigurationManager.AppSettings("PrivateKeyServer")
        Dim keyfile As String = Path.Combine(My.Application.Info.DirectoryPath, filename)
        Dim keyxml As String = File.ReadAllText(keyfile)
        Dim encrypted As String = RSAEncryption.Encrypt(cleartext, keyxml)
        Dim decrypted As String = RSAEncryption.Decrypt(encrypted, keyxml)
        Assert.IsTrue(decrypted = cleartext)
    End Sub


#Region "Create and Export RSA keys"

    <TestMethod()> Public Sub NewRSAKeyXMLTest()
        Dim xml As String = RSAEncryption.NewRSAKeyXML()
        Assert.IsTrue(xml.Length > 0)
    End Sub

    <TestMethod()> Public Sub PersistFetchDeleteRSAKeyXMLTest()
        Dim containername As String = "RSAKeyXMLTest"
        Dim xml As String = RSAEncryption.NewRSAKeyXML()
        RSAEncryption.PersistRSAKeyXML(xml, containername)
        Dim fetch As String = RSAEncryption.FetchRSAKeyXML(containername, False)
        Assert.IsTrue(fetch.Length > 0)
        RSAEncryption.DeleteRSAKeyXML(containername)
        Try
            Dim blank As String = RSAEncryption.FetchRSAKeyXML(containername, False)
            Assert.IsTrue(blank.Length = 0)
        Catch ex As Exception
            'We expect the second Fetch call to throw an error if the key has been deleted
            Assert.IsTrue(1 = 1)
        End Try
    End Sub

#End Region

#Region "RSA with XML Keys"

    <TestMethod()> Public Sub EncryptDecryptWithXMLKeysTest()
        Dim cleartext As String = "TEST_DATA_GOES_HERE"
        Dim xml As String = RSAEncryption.NewRSAKeyXML()
        Dim clear As Byte() = ASCIIEncoding.UTF8.GetBytes(cleartext)
        Dim encrypted As Byte() = RSAEncryption.Encrypt(clear, xml)
        Dim decrypted As Byte() = RSAEncryption.Decrypt(encrypted, xml)
        Dim plaintext As String = ASCIIEncoding.UTF8.GetString(decrypted)
        Assert.IsTrue(plaintext = cleartext)
    End Sub

#End Region

#Region "Create and Export RSA keys"

    <TestMethod()> Public Sub PersistFetchDeleteRSAKeyContainerTest()
        Dim containername As String = "RSAKeyXMLTest"
        Dim xml As String = RSAEncryption.NewRSAKeyXML()
        RSAEncryption.PersistRSAKeyContainer(containername)
        Dim cleartext As String = "TEST_DATA_GOES_HERE"
        Dim clearbytes As Byte() = ASCIIEncoding.UTF8.GetBytes(cleartext)
        Dim encryptedbytes As Byte() = RSAEncryption.Encrypt(clearbytes, containername, False)
        Assert.IsTrue(encryptedbytes.Length > 0)
        RSAEncryption.DeleteRSAKeyContainer(containername)
        Try
            Dim decryptedbytes As Byte() = RSAEncryption.Decrypt(encryptedbytes, containername, False)
            Assert.IsTrue(decryptedbytes.Length = 0)
        Catch ex As Exception
            'We expect the second Fetch call to throw an error if the key has been deleted
            Assert.IsTrue(1 = 1)
        End Try
    End Sub

#End Region

#Region "RSA with Keys in Containers"

    <TestMethod()> Public Sub EncryptDecryptWithContainersTest()
        Dim cleartext As String = "TEST_DATA_GOES_HERE"
        Dim containername As String = "EncryptDecryptWithContainersTest"
        RSAEncryption.PersistRSAKeyContainer(containername)
        Dim clearbytes As Byte() = ASCIIEncoding.UTF8.GetBytes(cleartext)
        Dim encryptedbytes As Byte() = RSAEncryption.Encrypt(clearbytes, containername, False)
        Dim decryptedbytes As Byte() = RSAEncryption.Decrypt(encryptedbytes, containername, False)
        Dim decrypted As String = ASCIIEncoding.UTF8.GetString(decryptedbytes)
        RSAEncryption.DeleteRSAKeyContainer(containername)
        Assert.IsTrue(decrypted = cleartext)
    End Sub

    <TestMethod()> Public Sub EncryptDecryptWithOtherPartyContainersWorkbench()
        Dim cleartext As String = "TEST_DATA_GOES_HERE"
        Dim privatekeycontainer As String = "privatekeycontainername"
        Dim publickeycontainer As String = "publickeycontainername"
        Dim encryptedbytes As Byte()
        Dim decryptedbytes As Byte()
        Dim decrypted As String

        Debug.Print("----- EncryptDecryptWithOtherPartyContainersWorkbench -----")

        'Create private key container
        Dim cspPrivate As New SSC.CspParameters
        cspPrivate.KeyContainerName = privatekeycontainer
        cspPrivate.Flags = Security.Cryptography.CspProviderFlags.UseDefaultKeyContainer
        Dim rsaPrivate As New SSC.RSACryptoServiceProvider(cspPrivate)
        rsaPrivate.PersistKeyInCsp = True
        Debug.Print("Private Key: " & vbCrLf & rsaPrivate.ToXmlString(True))

        'Export public part of key
        Dim xml As String
        xml = rsaPrivate.ToXmlString(False)
        Debug.Print("Private Key Public Export: " & vbCrLf & xml)
        'Store public key in new container
        Dim cspPublic As New SSC.CspParameters()
        cspPublic.KeyContainerName = publickeycontainer
        cspPublic.Flags = Security.Cryptography.CspProviderFlags.UseDefaultKeyContainer
        Dim rsaPublic As New SSC.RSACryptoServiceProvider(cspPublic)
        rsaPublic.FromXmlString(xml)
        rsaPublic.PersistKeyInCsp = True
        Debug.Print("Public Key: " & vbCrLf & rsaPrivate.ToXmlString(True))

        Dim clearbytes As Byte() = ASCIIEncoding.UTF8.GetBytes(cleartext)

        'Check that we can encrypt and decrypt correctly
        encryptedbytes = rsaPublic.Encrypt(clearbytes, False)
        decryptedbytes = rsaPrivate.Decrypt(encryptedbytes, False)
        decrypted = ASCIIEncoding.UTF8.GetString(decryptedbytes)
        Assert.IsTrue(decrypted = cleartext)

        decrypted = ""

        'Simulate separate processes by creating CSPs 
        '   from the keys stored in the containers
        Dim cspEncrypt As New SSC.CspParameters
        cspEncrypt.KeyContainerName = publickeycontainer 'Why is this not setting the right xml?
        cspEncrypt.Flags = Security.Cryptography.CspProviderFlags.UseDefaultKeyContainer
        Dim rsaEncryptor As New SSC.RSACryptoServiceProvider(cspEncrypt)
        'Debug.Print("Encryptor Key: " & vbCrLf & rsaEncryptor.ToXmlString(False))
        Debug.Print("Encryptor Key: " & vbCrLf & rsaEncryptor.ToXmlString(True))
        encryptedbytes = rsaEncryptor.Encrypt(clearbytes, False)

        'The two sets of encrypted data should be identical(?) as they are created using the same key
        'Assert.IsTrue(rsaPublic.Encrypt(clearbytes, False) Is rsaEncryptor.Encrypt(clearbytes, False))

        Dim cspDecrypt As New SSC.CspParameters
        cspDecrypt.KeyContainerName = privatekeycontainer
        cspDecrypt.Flags = Security.Cryptography.CspProviderFlags.UseDefaultKeyContainer
        Dim rsaDecryptor As New SSC.RSACryptoServiceProvider(cspDecrypt)
        Debug.Print("Decryptor Key: " & vbCrLf & rsaDecryptor.ToXmlString(True))
        decryptedbytes = rsaDecryptor.Decrypt(encryptedbytes, False)

        decrypted = ASCIIEncoding.UTF8.GetString(decryptedbytes)
        Assert.IsTrue(decrypted = cleartext)

        rsaPrivate.PersistKeyInCsp = False
        rsaPrivate.Clear()
        rsaPublic.PersistKeyInCsp = False
        rsaPublic.Clear()
        rsaEncryptor.PersistKeyInCsp = False
        rsaEncryptor.Clear()
        rsaDecryptor.PersistKeyInCsp = False
        rsaDecryptor.Clear()

    End Sub

    <TestMethod()> Public Sub EncryptDecryptWithOtherPartyContainersTest()
        Dim cleartext As String = "TEST_DATA_GOES_HERE"
        Dim privatekeycontainer As String = "privatekeycontainername"
        Dim publickeycontainer As String = "publickeycontainername"
        RSAEncryption.PersistRSAKeyContainer(privatekeycontainer)
        Dim xml As String = RSAEncryption.FetchRSAKeyPublicXML(privatekeycontainer)
        RSAEncryption.PersistRSAKeyXML(xml, publickeycontainer)
        Dim clearbytes As Byte() = ASCIIEncoding.UTF8.GetBytes(cleartext)
        Dim encryptedbytes As Byte() = RSAEncryption.Encrypt(clearbytes, publickeycontainer, False)
        Dim decryptedbytes As Byte() = RSAEncryption.Decrypt(encryptedbytes, privatekeycontainer, False)
        Dim decrypted As String = ASCIIEncoding.UTF8.GetString(decryptedbytes)
        RSAEncryption.DeleteRSAKeyContainer(privatekeycontainer)
        Assert.IsTrue(decrypted = cleartext)
    End Sub

#End Region

#Region "RSA with Parameters"

    <TestMethod()> Public Sub EncryptDecryptWithParametersTest()
        Dim cleartext As String = "TEST_DATA_GOES_HERE"

        Dim RSA As New SSC.RSACryptoServiceProvider
        Dim params As SSC.RSAParameters = RSA.ExportParameters(True)

        Dim clearbytes As Byte() = ASCIIEncoding.UTF8.GetBytes(cleartext)
        Dim encryptedbytes As Byte() = RSAEncryption.Encrypt(clearbytes, params, False)
        Dim decryptedbytes As Byte() = RSAEncryption.Decrypt(encryptedbytes, params, False)
        Dim decrypted As String = ASCIIEncoding.UTF8.GetString(decryptedbytes)

        Assert.IsTrue(decrypted = cleartext)

    End Sub

#End Region

End Class
