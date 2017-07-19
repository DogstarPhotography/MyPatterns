Imports Microsoft.Practices.EnterpriseLibrary.Data
Imports Microsoft.Practices.Unity
Imports Microsoft.Practices.EnterpriseLibrary.Common.Configuration.Unity

'http://msdn.microsoft.com/en-us/library/ff664719(v=pandp.50).aspx

'Usage
Public Class ExampleScenario

	Private db As Database

	Public Sub New(<Dependency("Sales")> ByVal theDatabase As Database)
		db = theDatabase

	End Sub

End Class

Public Class Whatever

	Public Sub stuff()

		' Resolve the class through the Unity container.
		Dim container = New UnityContainer().AddNewExtension(Of EnterpriseLibraryCoreExtension)()
		Dim myObject As ExampleScenario = container.Resolve(Of ExampleScenario)()

		'' Resolve the class through the default container. You must use this 
		'' approach if you are not using Unity with Enterprise Library. 
		'Dim myObject As ExampleScenario _
		'  = EnterpriseLibraryContainer.Current.GetInstance(Of ExampleScenario)()

	End Sub
End Class