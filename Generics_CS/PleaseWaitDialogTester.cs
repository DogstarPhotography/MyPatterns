//INSTANT C# NOTE: Formerly VB project-level imports:
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;

namespace Generics
{
	public class PleaseWaitDialogTester
	{

//INSTANT C# NOTE: These were formerly VB static local variables:
		private static int PleaseWaitTest_Counter = 0;

		public static void TestPleaseWaitDialog()
		{

			PleaseWaitDialog newDailog = new PleaseWaitDialog();
			//newDailog.Setup(ref PleaseWaitTest, 10000, 1000);

		}

		public static bool PleaseWaitTest()
		{
//INSTANT C# NOTE: VB local static variable moved to class level:
//			Static Counter As Integer = 0
			PleaseWaitTest_Counter = PleaseWaitTest_Counter + 1;
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