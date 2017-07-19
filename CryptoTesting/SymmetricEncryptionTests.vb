Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Crypto.Symmetric

<TestClass()> Public Class SymmetricEncryptionTests

#Region "AES"

    <TestMethod()> Public Sub EncryptDecryptAESTest()
        Dim data As String = "TEST_DATA_GOES_HERE"
        Dim password As String = "testpassword"
        Dim salt As String = "salt_salt_salt"
        Dim encrypted As String = AESEncryption.Encrypt(data, password, salt)
        Debug.Print(encrypted)
        Dim decrypted As String = AESEncryption.Decrypt(encrypted, password, salt)
        Assert.IsTrue(decrypted = data)
    End Sub

    <TestMethod()> Public Sub EncryptDecryptAES_BadParametersTest()
        Dim cleartext As String = ""
        Dim password As String = ""
        Dim salt As String = ""
        Dim encrypted As String
        Dim decrypted As String
        'Bad password
        Try
            encrypted = AESEncryption.Encrypt(cleartext, password, salt)
            Assert.Fail()
        Catch ex As Exception
            'We should get an exception for a bad password so correct the error
            password = "testpassword"
        End Try
        'Bad salt
        Try
            encrypted = AESEncryption.Encrypt(cleartext, password, salt)
            Assert.Fail()
        Catch ex As Exception
            'We should get an exception for a bad salt so correct the error
            salt = "salt_salt_salt"
        End Try
        'Blank cleartext
        Try
            encrypted = AESEncryption.Encrypt(cleartext, password, salt)
            Assert.Fail()
        Catch ex As Exception
            'We should get an exception for a blank cleartext so correct the error
            cleartext = "TEST_DATA_GOES_HERE"
        End Try
        encrypted = AESEncryption.Encrypt(cleartext, password, salt)
        Assert.IsTrue(encrypted.Length > 0)
        Debug.Print(encrypted)
        password = ""
        salt = ""
        'Bad password
        Try
            decrypted = AESEncryption.Decrypt(encrypted, password, salt)
            Assert.Fail()
        Catch ex As Exception
            'We should get an exception for a bad password so correct the error
            password = "testpassword"
        End Try
        'Bad salt
        Try
            decrypted = AESEncryption.Decrypt(encrypted, password, salt)
            Assert.Fail()
        Catch ex As Exception
            'We should get an exception for a bad salt so correct the error
            salt = "salt_salt_salt"
        End Try
        decrypted = AESEncryption.Decrypt(encrypted, password, salt)
        Assert.IsTrue(decrypted = cleartext)
    End Sub

    <TestMethod()> Public Sub EncryptDecryptAES_ComplexDataTest()
        Dim data As String = "<RSAKeyValue><Modulus>j3i5aBaEKtRiaGJI2TUqGT7xHNsT5P+dGy9lmrV7BK11iS2tSFubp2UPfpDYDMc77XJJkqUrdxNuYi7oqMZ/JgwaWWQDAngjmtpYFl8Nq7PsX9LV8yWlU6IZpBLLfFpT3z3J1+4t7LzIt0RyEX6t/TKD1lJNcnuHFYW9AokPbI8=</Modulus><Exponent>AQAB</Exponent><P>tzjPjEb3+oSSDOxivm5rQ38b1uVBGcXtTjrS+CvXDOPFJcRB2msXMLcrCGFUCOcn7rlDGOmZULZqlY+bULQc9w==</P><Q>yHXXlWbRPLcsnsfCjhRWAYmPvJgu4LfOovsavtEqPcy8Erv6A+4A6ozSyoRVQ2K74DkBNO89/6xkZTDBaBo/KQ==</Q><DP>M7ko9jvOo30rUdSlp4a6ZyqJ7Gd5slHqxQvcJM0Tf4MJU7kMsiFLQahj0JDRTVYcMstAAtdnPZ7RhfktamH+Tw==</DP><DQ>Z3VqbpFCLDPds5UltG6KdQCqTou8pf43h6ZRh2osgvjHmGOsBZswnd1QbXUfDEhI7tB87vUK6onuxssDBteFAQ==</DQ><InverseQ>j8QNGdbhUD85lXWwVa2IE1I9H/K1OyTbeMT+L50vnjMQc/h76zemy3RyH3SwLAttILcJnLiuReozDbNe/fVcTA==</InverseQ><D>BNxtx8GPi9XzWZ8O4dEj0oQn7jbcBzPD8nJaKnJAr0ljRJkYGG4GKZdKfZrRvykW9jYbmQzgmqG9aTU2q9VB5I6fbYRC1/RDhtnLszTSIRxNucM1sAgLyN90765PoirG+b4LHu8IYoXwdcoLnFIswxF1ErH083b0PTGjWsjiE3k=</D></RSAKeyValue>"
        Dim password As String = "testpassword"
        Dim salt As String = "salt_salt_salt"
        Dim encrypted As String = AESEncryption.Encrypt(data, password, salt)
        'Debug.Print(encrypted)
        Dim decrypted As String = AESEncryption.Decrypt(encrypted, password, salt)
        Assert.IsTrue(decrypted = data)
    End Sub

#End Region

#Region "AESHMAC"

    <TestMethod()> Public Sub EncryptDecryptAESHMACTest()
        Dim data As String = "TEST_DATA_GOES_HERE"
        Dim password As String = "testpassword"
        Dim encrypted As String = AESHMACEncryption.Encrypt(data, password)
        Debug.Print(encrypted)
        Dim decrypted As String = AESHMACEncryption.Decrypt(encrypted, password)
        Assert.IsTrue(decrypted = data)
    End Sub

    <TestMethod()> Public Sub EncryptDecryptAESHMAC_BadParametersTest()
        Dim cleartext As String = ""
        Dim password As String = ""
        Dim salt As String = ""
        Dim encrypted As String
        Dim decrypted As String
        'Bad password
        Try
            encrypted = AESHMACEncryption.Encrypt(cleartext, password)
            Assert.Fail()
        Catch ex As Exception
            'We should get an exception for a bad password so correct the error
            password = "testpassword"
        End Try
        'Blank cleartext
        Try
            encrypted = AESHMACEncryption.Encrypt(cleartext, password)
            Assert.Fail()
        Catch ex As Exception
            'We should get an exception for a blank cleartext so correct the error
            cleartext = "TEST_DATA_GOES_HERE"
        End Try
        encrypted = AESHMACEncryption.Encrypt(cleartext, password)
        Assert.IsTrue(encrypted.Length > 0)
        Debug.Print(encrypted)
        password = ""
        salt = ""
        'Bad password
        Try
            decrypted = AESHMACEncryption.Decrypt(encrypted, password)
            Assert.Fail()
        Catch ex As Exception
            'We should get an exception for a bad password so correct the error
            password = "testpassword"
        End Try
        decrypted = AESHMACEncryption.Decrypt(encrypted, password)
        Assert.IsTrue(decrypted = cleartext)
    End Sub

    <TestMethod()> Public Sub EncryptDecryptAESHMAC_ComplexDataTest()
        Dim data As String = "<RSAKeyValue><Modulus>j3i5aBaEKtRiaGJI2TUqGT7xHNsT5P+dGy9lmrV7BK11iS2tSFubp2UPfpDYDMc77XJJkqUrdxNuYi7oqMZ/JgwaWWQDAngjmtpYFl8Nq7PsX9LV8yWlU6IZpBLLfFpT3z3J1+4t7LzIt0RyEX6t/TKD1lJNcnuHFYW9AokPbI8=</Modulus><Exponent>AQAB</Exponent><P>tzjPjEb3+oSSDOxivm5rQ38b1uVBGcXtTjrS+CvXDOPFJcRB2msXMLcrCGFUCOcn7rlDGOmZULZqlY+bULQc9w==</P><Q>yHXXlWbRPLcsnsfCjhRWAYmPvJgu4LfOovsavtEqPcy8Erv6A+4A6ozSyoRVQ2K74DkBNO89/6xkZTDBaBo/KQ==</Q><DP>M7ko9jvOo30rUdSlp4a6ZyqJ7Gd5slHqxQvcJM0Tf4MJU7kMsiFLQahj0JDRTVYcMstAAtdnPZ7RhfktamH+Tw==</DP><DQ>Z3VqbpFCLDPds5UltG6KdQCqTou8pf43h6ZRh2osgvjHmGOsBZswnd1QbXUfDEhI7tB87vUK6onuxssDBteFAQ==</DQ><InverseQ>j8QNGdbhUD85lXWwVa2IE1I9H/K1OyTbeMT+L50vnjMQc/h76zemy3RyH3SwLAttILcJnLiuReozDbNe/fVcTA==</InverseQ><D>BNxtx8GPi9XzWZ8O4dEj0oQn7jbcBzPD8nJaKnJAr0ljRJkYGG4GKZdKfZrRvykW9jYbmQzgmqG9aTU2q9VB5I6fbYRC1/RDhtnLszTSIRxNucM1sAgLyN90765PoirG+b4LHu8IYoXwdcoLnFIswxF1ErH083b0PTGjWsjiE3k=</D></RSAKeyValue>"
        Dim password As String = "testpassword"
        Dim encrypted As String = AESHMACEncryption.Encrypt(data, password)
        Debug.Print(encrypted)
        Dim decrypted As String = AESHMACEncryption.Decrypt(encrypted, password)
        Assert.IsTrue(decrypted = data)
    End Sub

#End Region

#Region "Compare"

    <TestMethod()> Public Sub AESvsAESHMAC_EncryptedLengthTest()
        Dim data As String = "TEST_DATA_GOES_HERE_TEST_DATA_GOES_HERE_TEST_DATA_GOES_HERE"
        Dim password As String = "testpassword"
        Dim salt As String = "salt_salt_salt"
        Dim aes_encrypted As String = AESEncryption.Encrypt(data, password, salt)
        Dim aeshmac_encrypted As String = AESHMACEncryption.Encrypt(data, password)
        Debug.Print("    AES: " & aes_encrypted)
        Debug.Print("AESHMAC: " & aeshmac_encrypted)
        Assert.IsTrue(aes_encrypted.Length < aeshmac_encrypted.Length)
    End Sub

#End Region

End Class