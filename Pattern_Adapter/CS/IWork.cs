// VBConversions Note: VB project level imports
using System.Collections.Generic;
using System;
using System.Diagnostics;
using System.Data;
using Microsoft.VisualBasic;
using System.Collections;
// End of VB project level imports


/// <summary>
/// Simple interface for a class that does some work
/// </summary>
/// <remarks></remarks>
namespace Pattern_Adapter
{
	public interface IWork
	{
		/// <summary>
		/// Do some work
		/// </summary>
		/// <remarks></remarks>
		void DoWork();
	}
	
}
