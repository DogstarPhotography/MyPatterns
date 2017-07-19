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

using System.Data.SqlClient;
using System.Runtime.CompilerServices;


namespace EnterpriseLibrary_Patterns
{
	public static class BackingStoreExtensions
	{
		
		[Extension]public static string SafeGetString(this IDataReader reader, int colIndex)
		{
			if (reader.GetValue(colIndex) != DBNull.Value)
			{
				return reader.GetString(colIndex);
			}
			else
			{
				return string.Empty;
			}
		}
		
		//<Extension> Public Function SafeGetInt32(reader As IDataReader, colIndex As Integer) As Integer
		//    If reader.GetInt32(colIndex) IsNot DBNull.Value Then
		//        Return reader.GetString(colIndex)
		//    Else
		//        Return String.Empty
		//    End If
		//End Function
		
	}
	
}
