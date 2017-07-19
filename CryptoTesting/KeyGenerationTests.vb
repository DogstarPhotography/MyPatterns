Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Crypto.Symmetric
Imports Crypto

<TestClass()> Public Class KeyGenerationTests

    <TestMethod()> Public Sub GenerateKeyTest()
        Dim password As String = "testpassword"
        Dim salt As String = "salt_salt_salt"
        'Version 1
        Dim testA As String = ASCIIEncoding.UTF8.GetString(KeyGeneration.GenerateKey(password, salt, 32))
        Debug.Print(testA)
        Dim testB As String = ASCIIEncoding.UTF8.GetString(KeyGeneration.GenerateKey(password, salt, 32))
        Debug.Print(testB)
        Assert.IsTrue(testA = testB)
        Dim testC As String = ASCIIEncoding.UTF8.GetString(KeyGeneration.GenerateKey(password, salt, 16))
        Debug.Print(testC)
        Dim testD As String = ASCIIEncoding.UTF8.GetString(KeyGeneration.GenerateKey(password, salt, 16))
        Debug.Print(testD)
        Assert.IsTrue(testC = testD)
        'Version 2
        Dim saltbytes As Byte() = Encoding.ASCII.GetBytes(salt)
        Dim testE As String = ASCIIEncoding.UTF8.GetString(KeyGeneration.GenerateKey(password, saltbytes))
        Debug.Print(testA)
        Dim testF As String = ASCIIEncoding.UTF8.GetString(KeyGeneration.GenerateKey(password, saltbytes))
        Debug.Print(testB)
        Assert.IsTrue(testE = testF)
    End Sub

    <TestMethod()> Public Sub GenerateKeySaltPairTest()
        Dim password As String = "test_password"
        Dim testA As KeySaltPair = KeyGeneration.GenerateKeySaltPair(password)
        Debug.Print(ASCIIEncoding.UTF8.GetString(testA.Key))
        Debug.Print(ASCIIEncoding.UTF8.GetString(testA.Salt))
        Dim testB As KeySaltPair = KeyGeneration.GenerateKeySaltPair(password)
        Debug.Print(ASCIIEncoding.UTF8.GetString(testB.Key))
        Debug.Print(ASCIIEncoding.UTF8.GetString(testB.Salt))
        Assert.IsTrue(ASCIIEncoding.UTF8.GetString(testA.Key) <> ASCIIEncoding.UTF8.GetString(testB.Key))
        Assert.IsTrue(ASCIIEncoding.UTF8.GetString(testA.Salt) <> ASCIIEncoding.UTF8.GetString(testB.Salt))
    End Sub

    <TestMethod()> Public Sub GenerateKeySaltPairsTest()
        Dim password As String = "test_password"
        Dim pairs As List(Of KeySaltPair) = KeyGeneration.GenerateKeySaltPair(password, 2)
        Dim testA As KeySaltPair = pairs.Item(0)
        Debug.Print(ASCIIEncoding.UTF8.GetString(testA.Key))
        Debug.Print(ASCIIEncoding.UTF8.GetString(testA.Salt))
        Assert.IsTrue(testA.Key.Length = EncryptionSettings.KeyBytes)
        Dim testB As KeySaltPair = pairs.Item(1)
        Debug.Print(ASCIIEncoding.UTF8.GetString(testB.Key))
        Debug.Print(ASCIIEncoding.UTF8.GetString(testB.Salt))
        Assert.IsTrue(testB.Key.Length = EncryptionSettings.KeyBytes)
        Assert.IsTrue(ASCIIEncoding.UTF8.GetString(testA.Key) <> ASCIIEncoding.UTF8.GetString(testB.Key))
        Assert.IsTrue(ASCIIEncoding.UTF8.GetString(testA.Salt) = ASCIIEncoding.UTF8.GetString(testB.Salt))
    End Sub

End Class