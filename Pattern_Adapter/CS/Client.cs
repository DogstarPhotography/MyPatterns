// VBConversions Note: VB project level imports
using System.Collections.Generic;
using System;
using System.Diagnostics;
using System.Data;
using Microsoft.VisualBasic;
using System.Collections;
// End of VB project level imports


namespace Pattern_Adapter
{
	public class Client
	{
		
		public void UseAdapter()
		{
			
			ITarget MyAdapter = new ObjectAdapter();
			
			MyAdapter.Request();
			
		}
	}
	
}
