// VBConversions Note: VB project level imports
using System.Collections.Generic;
using System;
using System.Linq;
using System.Drawing;
using System.Diagnostics;
using System.Data;
using System.Xml.Linq;
using Microsoft.VisualBasic;
using System.Collections;
using System.Windows.Forms;
// End of VB project level imports

using Microsoft.Practices.EnterpriseLibrary.Data;
using Microsoft.Practices.Unity;
using Microsoft.Practices.EnterpriseLibrary.Common.Configuration.Unity;


//http://msdn.microsoft.com/en-us/library/ff664719(v=pandp.50).aspx

//Usage
namespace EnterpriseLibrary_Patterns
{
	public class ExampleScenario
	{
		
		private Microsoft.Practices.EnterpriseLibrary.Data.Database db;
		
		public ExampleScenario(Microsoft.Practices.EnterpriseLibrary.Data.Database theDatabase)
		{
			db = theDatabase;
			
		}
		
	}
	
	public class Whatever
	{
		
		public void stuff()
		{
			
			// Resolve the class through the Unity container.
			var container = new UnityContainer().AddNewExtension<EnterpriseLibraryCoreExtension>();
			ExampleScenario myObject = container.Resolve<ExampleScenario>();
			
			//' Resolve the class through the default container. You must use this
			//' approach if you are not using Unity with Enterprise Library.
			//Dim myObject As ExampleScenario _
			//  = EnterpriseLibraryContainer.Current.GetInstance(Of ExampleScenario)()
			
		}
	}
}
