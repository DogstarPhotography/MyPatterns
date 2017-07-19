//INSTANT C# NOTE: Formerly VB project-level imports:
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;

namespace Generics
{
	public partial class PleaseWaitDialog
	{
		/// <summary>
		/// Delegate for callback function
		/// </summary>
		/// <returns></returns>
		/// <remarks></remarks>
		internal PleaseWaitDialog()
		{
			InitializeComponent();
		}

		public delegate bool PleaseWaitCallback();

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
		public void Setup(ref PleaseWaitCallback CallbackFunction, ushort TimeoutMilliseconds, ushort IntervalMilliseconds)
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
			catch (Exception ex)
			{
				//Ignore
			}
		}

		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			try
			{
				this.Close();
			}
			catch (Exception ex)
			{
				//Ignore
			}
		}

		private void picDistraction_Click(object sender, System.EventArgs e)
		{
			try
			{
				this.Close();
			}
			catch (Exception ex)
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
		private void myTimer_Tick(object sender, System.EventArgs e)
		{
			try
			{
				bool Result = myCallbackDelegate.Invoke();
				myTimeout = myTimeout - myTimerCycle;
				if (Result == true || myTimeout < 0)
				{
					myTimer.Enabled = false;
					this.Close();
				}
			}
			catch (Exception ex)
			{
				//Ignore
			}
		}
	}
}