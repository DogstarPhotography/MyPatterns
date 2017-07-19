//INSTANT C# NOTE: Formerly VB project-level imports:
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;

namespace Generics
{
	[Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
	public partial class PleaseWaitDialog : System.Windows.Forms.Form
	{
		//Form overrides dispose to clean up the component list.
		[System.Diagnostics.DebuggerNonUserCode()]
		protected override void Dispose(bool disposing)
		{
			try
			{
				if (disposing && components != null)
				{
					components.Dispose();
				}
			}
			finally
			{
				base.Dispose(disposing);
			}
		}

		//Required by the Windows Form Designer
		private System.ComponentModel.IContainer components;

		//NOTE: The following procedure is required by the Windows Form Designer
		//It can be modified using the Windows Form Designer.  
		//Do not modify it using the code editor.
		[System.Diagnostics.DebuggerStepThrough()]
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PleaseWaitDialog));
			this.picDistraction = new System.Windows.Forms.PictureBox();
			this.btnCancel = new System.Windows.Forms.Button();
			this.myTimer = new System.Windows.Forms.Timer(this.components);
			((System.ComponentModel.ISupportInitialize)this.picDistraction).BeginInit();
			this.SuspendLayout();
			//
			//picDistraction
			//
			this.picDistraction.Anchor = (System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) | System.Windows.Forms.AnchorStyles.Left) | System.Windows.Forms.AnchorStyles.Right);
			this.picDistraction.ErrorImage = (System.Drawing.Image)resources.GetObject("picDistraction.ErrorImage");
			this.picDistraction.Image = (System.Drawing.Image)resources.GetObject("picDistraction.Image");
			this.picDistraction.InitialImage = (System.Drawing.Image)resources.GetObject("picDistraction.InitialImage");
			this.picDistraction.Location = new System.Drawing.Point(12, 12);
			this.picDistraction.Name = "picDistraction";
			this.picDistraction.Size = new System.Drawing.Size(175, 100);
			this.picDistraction.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
			this.picDistraction.TabIndex = 0;
			this.picDistraction.TabStop = false;
			//
			//btnCancel
			//
			this.btnCancel.Anchor = (System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) | System.Windows.Forms.AnchorStyles.Right);
			this.btnCancel.Location = new System.Drawing.Point(12, 118);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(175, 30);
			this.btnCancel.TabIndex = 1;
			this.btnCancel.Text = "Cancel";
			this.btnCancel.UseVisualStyleBackColor = true;
			//
			//PleaseWaitDialog
			//
			this.AutoScaleDimensions = new System.Drawing.SizeF(6.0F, 13.0F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(199, 160);
			this.Controls.Add(this.btnCancel);
			this.Controls.Add(this.picDistraction);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "PleaseWaitDialog";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Please Wait";
			((System.ComponentModel.ISupportInitialize)this.picDistraction).EndInit();
			this.ResumeLayout(false);

//INSTANT C# NOTE: Converted design-time event handler wireups:
			btnCancel.Click += new System.EventHandler(btnCancel_Click);
			picDistraction.Click += new System.EventHandler(picDistraction_Click);
			myTimer.Tick += new System.EventHandler(myTimer_Tick);
		}
		internal System.Windows.Forms.PictureBox picDistraction;
		internal System.Windows.Forms.Button btnCancel;
		internal System.Windows.Forms.Timer myTimer;
	}

}