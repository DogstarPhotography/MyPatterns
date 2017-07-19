Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Crypto

<TestClass()> Public Class RandomiserTests

    <TestMethod()> Public Sub ShuffleTest()
        Dim text As String = "thisisnotagoodpassword"
        Dim result As String = Randomiser.Shuffle(text)
        Debug.Print(result)
        Assert.IsTrue(result <> text)
    End Sub

    <TestMethod()> Public Sub GenerateTest()
        Dim text As String = "thisisnotagoodpassword"
        Dim result As String
        result = Randomiser.Generate()
        Debug.Print(result)
        Assert.IsTrue(result <> text)
        Dim key As String = "thisisnotagoodkey"
        result = Randomiser.Generate(key)
        Debug.Print(result)
        Assert.IsTrue(result <> text)
        Dim salt As String = "thisisnotagoodsalt"
        result = Randomiser.Generate(key, salt)
        Debug.Print(result)
        Assert.IsTrue(result <> text)
        Dim length As Integer = 12
        result = Randomiser.Generate(key, salt, length)
        Debug.Print(result)
        Assert.IsTrue(result <> text)
        Assert.IsTrue(result.Length = length)
    End Sub

End Class