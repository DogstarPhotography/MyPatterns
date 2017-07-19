Option Strict Off 'Required for LateBindingTest as the method name are used in the source code

Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports System.Reflection
Imports System.IO
''' <summary>
''' 
''' </summary>
''' <remarks>
''' Note that the target assembly must be in the same folder as the unit test - for LateBinding tests both project are built into the LateBnding project folder
''' </remarks>
<TestClass()> Public Class LateBindingTests

    <TestMethod()> Public Sub LateBindingWrapperTest()
        Dim instance As Object = LateBindingWrapper.GetInstance()
        Dim testdata As String = "TEST_DATA"
        LateBindingWrapper.SetData(instance, testdata)
        Dim result As String = LateBindingWrapper.GetData(instance)
        Assert.IsTrue(result.Length > 0)
        Assert.IsTrue(result = testdata)
    End Sub

    <TestMethod()> Public Sub LateBindingSingletonWrapperTest()
        Dim testdata As String = "TEST_DATA"
        LateBindingSingletonWrapper.Instance.SetData(testdata)
        Dim result As String = LateBindingSingletonWrapper.Instance.GetData()
        Assert.IsTrue(result.Length > 0)
        Assert.IsTrue(result = testdata)
    End Sub

    <TestMethod()> Public Sub LateBindingTest()
        Dim a As Assembly = Nothing
        Try
            a = Assembly.Load("MyPatterns.LateBinding")
        Catch e As FileNotFoundException
            Assert.Fail()
        End Try
        Dim classType As Type = a.[GetType]("MyPatterns.LateBinding.LateBound")
        Dim instance As Object = Activator.CreateInstance(classType)
        Assert.IsTrue(instance IsNot Nothing)
        Dim testdata As String = "TEST_DATA"
        instance.SetData(testdata)
        Dim result As String = instance.GetData()
        Assert.IsTrue(result.Length > 0)
        Assert.IsTrue(result = testdata)
    End Sub

    <TestMethod()> Public Sub InvokeMemberTest()
        Dim a As Assembly = Nothing
        Try
            a = Assembly.Load("MyPatterns.LateBinding")
        Catch e As FileNotFoundException
            Assert.Fail()
        End Try
        Dim classType As Type = a.[GetType]("MyPatterns.LateBinding.LateBound")
        Dim instance As Object = Activator.CreateInstance(classType)
        Assert.IsTrue(instance IsNot Nothing)
        Dim testdata As String = "TEST_DATA"
        classType.InvokeMember("SetData", BindingFlags.InvokeMethod, Nothing, instance, {testdata})
        Dim result As String = CType(classType.InvokeMember("GetData", BindingFlags.InvokeMethod, Nothing, instance, Nothing), String)
        Assert.IsTrue(result.Length > 0)
        Assert.IsTrue(result = testdata)
    End Sub

End Class