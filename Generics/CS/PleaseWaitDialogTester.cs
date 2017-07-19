// VBConversions Note: VB project level imports
using System.Collections.Generic;
using System;
using System.Diagnostics;
using System.Data;
using Microsoft.VisualBasic;
using System.Collections;
// End of VB project level imports


namespace Generics
{
	public class PleaseWaitDialogTester
	{
		
		public static void TestPleaseWaitDialog()
		{
			
			PleaseWaitDialog newDailog = new PleaseWaitDialog();
			newDailog.Setup(PleaseWaitTest, 10000, 1000);
			
		}
		
		// VBConversions Note: Former VB static variables moved to class level because they aren't supported in C#.
		static int PleaseWaitTest_Counter = 0;
		
		public static bool PleaseWaitTest()
		{
			// static int Counter = 0; VBConversions Note: Static variable moved to class level and renamed PleaseWaitTest_Counter. Local static variables are not supported in C#.
			PleaseWaitTest_Counter++;
			if (PleaseWaitTest_Counter > 7)
			{
				return true;
			}
			else
			{
				return false;
			}
		}
	}
	
}
