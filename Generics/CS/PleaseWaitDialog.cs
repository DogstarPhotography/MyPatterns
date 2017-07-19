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
	public partial class PleaseWaitDialog
	{
		public PleaseWaitDialog()
		{
			InitializeComponent();
		}
		/// <summary>
		/// Delegate for callback function
		/// </summary>
		/// <returns></returns>
		/// <remarks></remarks>
		public delegate bool  PleaseWaitCallback();
		
		private PleaseWaitCallback myCallbackDelegate;
		private int myTimeout;
		private int myTimerCycle;
		/// <summary>
		/// Shows the please wait dialog while calling the callback function regularly as determined by the IntervalMilliseconds.
		/// Will close when the callback function returns true or after approximately TimeoutMilliseconds have elapsed
		/// </summary>
		/// <param name="CallbackFunction"></param>
		/// <param name="TimeoutMilliseconds"></param>
		/// <param name="IntervalMilliseconds"></param>
		/// <remarks></remarks>
		public void Setup(PleaseWaitCallback CallbackFunction, UInt16 TimeoutMilliseconds, UInt16 IntervalMilliseconds)
		{
			try
			{
				myCallbackDelegate = CallbackFunction;
				myTimeout = TimeoutMilliseconds;
				myTimerCycle = IntervalMilliseconds;
				//Fire up the timer
				myTimer.Interval = myTimerCycle;
				myTimer.Enabled = true;
				//showdialog - this will block until the timer or the form calls close
				this.ShowDialog();
			}
			catch (Exception)
			{
				//Ignore
			}
		}
		
		public void btnCancel_Click(System.Object sender, System.EventArgs e)
		{
			try
			{
				this.Close();
			}
			catch (Exception)
			{
				//Ignore
			}
		}
		
		public void picDistraction_Click(System.Object sender, System.EventArgs e)
		{
			try
			{
				this.Close();
			}
			catch (Exception)
			{
				//Ignore
			}
		}
		/// <summary>
		/// Every time this comes round invoke the callback function and check the result, if the result is true or we have run out of time close the dialog
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void myTimer_Tick(object sender, System.EventArgs e)
		{
			try
			{
				bool Result = System.Convert.ToBoolean(myCallbackDelegate.Invoke());
				myTimeout = myTimeout - myTimerCycle;
				if (Result == true || myTimeout < 0)
				{
					myTimer.Enabled = false;
					this.Close();
				}
			}
			catch (Exception)
			{
				//Ignore
			}
		}
	}
}
