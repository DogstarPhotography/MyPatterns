Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Pattern_XMLSerialization

<TestClass()> Public Class SerializationTests

    <TestMethod()> Public Sub XMLAttributeExamples_SerializingListsTest()
        Dim target As New ListSerializationDemo
        Dim xml As String = target.Serialize()
        Debug.Print(xml)
        Assert.IsTrue(xml.Length > 0)
    End Sub

End Class