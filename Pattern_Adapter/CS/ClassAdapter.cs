// VBConversions Note: VB project level imports
using System.Collections.Generic;
using System;
using System.Diagnostics;
using System.Data;
using Microsoft.VisualBasic;
using System.Collections;
// End of VB project level imports


/// <summary>
///
/// </summary>
/// <remarks></remarks>
namespace Pattern_Adapter
{
	public class ClassAdapter : ITarget, IWork
	{
		//Private instance of adaptee
		private IWork MyAdaptee = new ClassAdaptee();
		/// <summary>
		/// Adapt ITarget.Request into Adaptee.Use
		/// </summary>
		/// <remarks></remarks>
		public void Request()
		{
			DoWork();
		}
		
		public void DoWork()
		{
			MyAdaptee.DoWork();
		}
	}
	
}
