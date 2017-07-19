Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Crypto

<TestClass()> Public Class HashTests

    <TestMethod()> Public Sub PBKDF2Test()
        Dim text As String = "PASSWORDHERE" '12 chars
        Dim salt As String = "SALTSALTSALT" '12 chars
        Dim length As Integer = 24
        Dim hashed As String = Hash.PBKDF2(text, salt, length)
        Assert.IsTrue(hashed.Length > 0)
        Debug.Print(hashed)
        Assert.IsTrue(hashed.Length = length)
        Dim rehashed As String = Hash.PBKDF2(text, salt, length)
        Debug.Print(rehashed)
        Assert.IsTrue(hashed = rehashed)
    End Sub

    <TestMethod()> Public Sub SHA256Test()
        Dim data As String = "TEST_DATA_GOES_HERE"
        Dim hashed As String = Hash.SHA256(data)
        Assert.IsTrue(hashed.Length > 0)
        Debug.Print(hashed)
        Dim rehashed As String = Hash.SHA256(data)
        Debug.Print(rehashed)
        Assert.IsTrue(hashed = rehashed)
        Assert.IsTrue(hashed.Length = 44)
    End Sub

    <TestMethod()> Public Sub SHA512Test()
        Dim data As String = "TEST_DATA_GOES_HERE"
        Dim hashed As String = Hash.SHA512(data)
        Assert.IsTrue(hashed.Length > 0)
        Debug.Print(hashed)
        Dim rehashed As String = Hash.SHA512(data)
        Debug.Print(rehashed)
        Assert.IsTrue(hashed = rehashed)
        Assert.IsTrue(hashed.Length = 88)
    End Sub

End Class