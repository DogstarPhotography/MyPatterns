Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports EnterpriseLibrary_Patterns
Imports EnterpriseLibrary_Patterns.Datafile
Imports System.Text

<TestClass()>
Public Class EnterpriseLibraryPatterns_DatafilesTest

	Private testContextInstance As TestContext

	'''<summary>
	'''Gets or sets the test context which provides
	'''information about and functionality for the current test run.
	'''</summary>
	Public Property TestContext() As TestContext
		Get
			Return testContextInstance
		End Get
		Set(ByVal value As TestContext)
			testContextInstance = value
		End Set
	End Property

#Region "Additional test attributes"
	'
	' You can use the following additional attributes as you write your tests:
	'
	' Use ClassInitialize to run code before running the first test in the class
	' <ClassInitialize()> Public Shared Sub MyClassInitialize(ByVal testContext As TestContext)
	' End Sub
	'
	' Use ClassCleanup to run code after all tests in a class have run
	' <ClassCleanup()> Public Shared Sub MyClassCleanup()
	' End Sub
	'
	' Use TestInitialize to run code before running each test
	' <TestInitialize()> Public Sub MyTestInitialize()
	' End Sub
	'
	' Use TestCleanup to run code after each test has run
	' <TestCleanup()> Public Sub MyTestCleanup()
	' End Sub
	'
#End Region

	'''<summary>
	'''A test for AddDatafile
	'''</summary>
	<TestMethod()> _
	Public Sub AddFileTest()
		Dim target As Datafiles = New Datafiles()
		Dim filepath As String = IO.Path.Combine(Environment.CurrentDirectory, "TestData.xml")
		target.AddFile(filepath)
		Assert.IsTrue(target.Count > 0)
	End Sub
	'''<summary>
	'''A test for LoadDatafiles
	'''</summary>
	<TestMethod()> _
	Public Sub LoadDatafilesTest()
		Dim target As Datafiles = New Datafiles()
		target.LoadDatafiles()
		Assert.IsTrue(target.Count > 0)
	End Sub


	'''<summary>
	'''A test for GetFileID
	'''</summary>
	<TestMethod()> _
	Public Sub GetFileIDTest()
		Dim target As Datafiles = New Datafiles()
		target.LoadDatafiles()
		Dim folder As String = "Default"
		Dim filename As String = "TestData"
		Dim extension As String = "xml"
		Dim expected As Integer = 5
		Dim actual As Integer
		actual = target.GetFileID(folder, filename, extension)
		Assert.AreEqual(expected, actual)
	End Sub

	'''<summary>
	'''A test for GetBinary
	'''</summary>
	<TestMethod()> _
	Public Sub GetBinaryTest()
		Dim target As Datafiles = New Datafiles()
		target.LoadDatafiles()
		Dim folder As String = "Default"
		Dim filename As String = "TestData"
		Dim extension As String = "xml"
		Dim id As Integer
		id = target.GetFileID(folder, filename, extension)
		Dim actual() As Byte
		actual = target.GetBinary(id)
		'Write out the actual bytes with their ASCII characters
		'For x = 0 To actual.Length - 1
		'	TestContext.WriteLine(actual(x) & " = " & Encoding.ASCII.GetChars(actual, x, 1))
		'Next
		Assert.IsTrue(actual.Length > 0)
	End Sub

	'''<summary>
	'''A test for GetText
	'''</summary>
	<TestMethod()> _
	Public Sub GetTextTest()
		Dim target As Datafiles = New Datafiles()
		target.LoadDatafiles()
		Dim folder As String = "Default"
		Dim filename As String = "TestData"
		Dim extension As String = "xml"
		Dim id As Integer
		id = target.GetFileID(folder, filename, extension)
		Dim actual As String
		actual = target.GetText(id)
		TestContext.WriteLine(actual)
		Assert.IsTrue(actual.Length > 0)
	End Sub
End Class
