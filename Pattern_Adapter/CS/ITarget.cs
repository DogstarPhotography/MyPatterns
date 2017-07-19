// VBConversions Note: VB project level imports
using System.Collections.Generic;
using System;
using System.Diagnostics;
using System.Data;
using Microsoft.VisualBasic;
using System.Collections;
// End of VB project level imports


/// <summary>
/// Common interface for Client and Adapter to use
/// </summary>
/// <remarks></remarks>
namespace Pattern_Adapter
{
	public interface ITarget
	{
		/// <summary>
		///
		/// </summary>
		/// <remarks></remarks>
		void Request();
		
	}
	
}
