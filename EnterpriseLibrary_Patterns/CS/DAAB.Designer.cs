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


namespace EnterpriseLibrary_Patterns
{
	[global::Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]public 
	partial class DAAB : System.Windows.Forms.Form
	{
		
		//Form overrides dispose to clean up the component list.
		[System.Diagnostics.DebuggerNonUserCode()]protected override void Dispose(bool disposing)
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
		private System.ComponentModel.Container components = null;
		
		//NOTE: The following procedure is required by the Windows Form Designer
		//It can be modified using the Windows Form Designer.
		//Do not modify it using the code editor.
		[System.Diagnostics.DebuggerStepThrough()]private void InitializeComponent()
		{
			this.dgvSampleData = new System.Windows.Forms.DataGridView();
			base.Load += DAAB_Load;
			this.GroupBox1 = new System.Windows.Forms.GroupBox();
			this.btnAccessor = new System.Windows.Forms.Button();
			this.btnAccessor.Click += this.btnAccessor_Click;
			this.btnDelete = new System.Windows.Forms.Button();
			this.btnDelete.Click += this.btnDelete_Click;
			this.btnOutParameter = new System.Windows.Forms.Button();
			this.btnOutParameter.Click += this.btnOutParameter_Click;
			this.btnSingleRow = new System.Windows.Forms.Button();
			this.btnSingleRow.Click += this.btnSingleRow_Click;
			this.btnInsert = new System.Windows.Forms.Button();
			this.btnInsert.Click += this.btnInsert_Click;
			this.btnScalar = new System.Windows.Forms.Button();
			this.btnScalar.Click += this.btnScalar_Click;
			this.btnDataSetToGrid = new System.Windows.Forms.Button();
			this.btnDataSetToGrid.Click += this.btnDataSetToGrid_Click;
			this.btnReader = new System.Windows.Forms.Button();
			this.btnReader.Click += this.btnReader_Click;
			this.txtResults = new System.Windows.Forms.TextBox();
			((System.ComponentModel.ISupportInitialize) this.dgvSampleData).BeginInit();
			this.GroupBox1.SuspendLayout();
			this.SuspendLayout();
			//
			//dgvSampleData
			//
			this.dgvSampleData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.dgvSampleData.Location = new System.Drawing.Point(358, 12);
			this.dgvSampleData.Name = "dgvSampleData";
			this.dgvSampleData.Size = new System.Drawing.Size(611, 279);
			this.dgvSampleData.TabIndex = 0;
			//
			//GroupBox1
			//
			this.GroupBox1.Controls.Add(this.btnAccessor);
			this.GroupBox1.Controls.Add(this.btnDelete);
			this.GroupBox1.Controls.Add(this.btnOutParameter);
			this.GroupBox1.Controls.Add(this.btnSingleRow);
			this.GroupBox1.Controls.Add(this.btnInsert);
			this.GroupBox1.Controls.Add(this.btnScalar);
			this.GroupBox1.Controls.Add(this.btnDataSetToGrid);
			this.GroupBox1.Controls.Add(this.btnReader);
			this.GroupBox1.Location = new System.Drawing.Point(12, 12);
			this.GroupBox1.Name = "GroupBox1";
			this.GroupBox1.Size = new System.Drawing.Size(340, 279);
			this.GroupBox1.TabIndex = 1;
			this.GroupBox1.TabStop = false;
			this.GroupBox1.Text = "Examples";
			//
			//btnAccessor
			//
			this.btnAccessor.Location = new System.Drawing.Point(6, 222);
			this.btnAccessor.Name = "btnAccessor";
			this.btnAccessor.Size = new System.Drawing.Size(328, 23);
			this.btnAccessor.TabIndex = 6;
			this.btnAccessor.Text = "Accessor";
			this.btnAccessor.UseVisualStyleBackColor = true;
			//
			//btnDelete
			//
			this.btnDelete.Location = new System.Drawing.Point(6, 193);
			this.btnDelete.Name = "btnDelete";
			this.btnDelete.Size = new System.Drawing.Size(328, 23);
			this.btnDelete.TabIndex = 6;
			this.btnDelete.Text = "Delete";
			this.btnDelete.UseVisualStyleBackColor = true;
			//
			//btnOutParameter
			//
			this.btnOutParameter.Location = new System.Drawing.Point(6, 164);
			this.btnOutParameter.Name = "btnOutParameter";
			this.btnOutParameter.Size = new System.Drawing.Size(328, 23);
			this.btnOutParameter.TabIndex = 5;
			this.btnOutParameter.Text = "Get Out Parameter";
			this.btnOutParameter.UseVisualStyleBackColor = true;
			//
			//btnSingleRow
			//
			this.btnSingleRow.Location = new System.Drawing.Point(6, 135);
			this.btnSingleRow.Name = "btnSingleRow";
			this.btnSingleRow.Size = new System.Drawing.Size(328, 23);
			this.btnSingleRow.TabIndex = 4;
			this.btnSingleRow.Text = "Get Single Row";
			this.btnSingleRow.UseVisualStyleBackColor = true;
			//
			//btnInsert
			//
			this.btnInsert.Location = new System.Drawing.Point(6, 106);
			this.btnInsert.Name = "btnInsert";
			this.btnInsert.Size = new System.Drawing.Size(328, 23);
			this.btnInsert.TabIndex = 3;
			this.btnInsert.Text = "GetStoredProccommand Insert";
			this.btnInsert.UseVisualStyleBackColor = true;
			//
			//btnScalar
			//
			this.btnScalar.Location = new System.Drawing.Point(6, 19);
			this.btnScalar.Name = "btnScalar";
			this.btnScalar.Size = new System.Drawing.Size(328, 23);
			this.btnScalar.TabIndex = 2;
			this.btnScalar.Text = "ExecuteScalar";
			this.btnScalar.UseVisualStyleBackColor = true;
			//
			//btnDataSetToGrid
			//
			this.btnDataSetToGrid.Location = new System.Drawing.Point(6, 48);
			this.btnDataSetToGrid.Name = "btnDataSetToGrid";
			this.btnDataSetToGrid.Size = new System.Drawing.Size(328, 23);
			this.btnDataSetToGrid.TabIndex = 1;
			this.btnDataSetToGrid.Text = "DataSet Grid Bind";
			this.btnDataSetToGrid.UseVisualStyleBackColor = true;
			//
			//btnReader
			//
			this.btnReader.Location = new System.Drawing.Point(6, 77);
			this.btnReader.Name = "btnReader";
			this.btnReader.Size = new System.Drawing.Size(328, 23);
			this.btnReader.TabIndex = 0;
			this.btnReader.Text = "DataReader";
			this.btnReader.UseVisualStyleBackColor = true;
			//
			//txtResults
			//
			this.txtResults.Location = new System.Drawing.Point(12, 297);
			this.txtResults.Multiline = true;
			this.txtResults.Name = "txtResults";
			this.txtResults.ScrollBars = System.Windows.Forms.ScrollBars.Both;
			this.txtResults.Size = new System.Drawing.Size(957, 321);
			this.txtResults.TabIndex = 3;
			//
			//DAAB
			//
			this.AutoScaleDimensions = new System.Drawing.SizeF((float) (6.0F), (float) (13.0F));
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(982, 630);
			this.Controls.Add(this.txtResults);
			this.Controls.Add(this.GroupBox1);
			this.Controls.Add(this.dgvSampleData);
			this.Name = "DAAB";
			this.Text = "DAAB";
			((System.ComponentModel.ISupportInitialize) this.dgvSampleData).EndInit();
			this.GroupBox1.ResumeLayout(false);
			this.ResumeLayout(false);
			this.PerformLayout();
			
		}
		internal System.Windows.Forms.DataGridView dgvSampleData;
		internal System.Windows.Forms.GroupBox GroupBox1;
		internal System.Windows.Forms.Button btnReader;
		internal System.Windows.Forms.TextBox txtResults;
		internal System.Windows.Forms.Button btnDataSetToGrid;
		internal System.Windows.Forms.Button btnScalar;
		internal System.Windows.Forms.Button btnInsert;
		internal System.Windows.Forms.Button btnSingleRow;
		internal System.Windows.Forms.Button btnOutParameter;
		internal System.Windows.Forms.Button btnDelete;
		internal System.Windows.Forms.Button btnAccessor;
	}
	
}
