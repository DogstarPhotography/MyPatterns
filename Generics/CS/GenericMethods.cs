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
	public class GenericMethods
	{
		/// <summary>
		/// Convert a list of one type to an array of another type
		/// </summary>
		/// <typeparam name="L"></typeparam>
		/// <typeparam name="A"></typeparam>
		/// <param name="Items"></param>
		/// <param name="Conversion"></param>
		/// <returns></returns>
		/// <remarks></remarks>
		public static A[] ConvertListToArray<L, A>(System.Collections.Generic.List<L> Items, Converter<L, A> Conversion)
		{
			
			//Use converter delegate to convert list into another list
			System.Collections.Generic.List<A> ConvertedList = Items.ConvertAll(Conversion);
			//Create an array to hold the data
			A[] ReturnArray = new A[ConvertedList.Count - 1 + 1];
			//Copy the list to the array
			ConvertedList.CopyTo(ReturnArray);
			//All done
			return ReturnArray;
			
		}
		/// <summary>
		/// Convert an array of one type to a list of another type
		/// </summary>
		/// <typeparam name="A"></typeparam>
		/// <typeparam name="L"></typeparam>
		/// <param name="Items"></param>
		/// <param name="Conversion"></param>
		/// <returns></returns>
		/// <remarks></remarks>
		public static System.Collections.Generic.List<L> ConvertArrayToList<A, L>(A[] Items, Converter<A, L> Conversion)
		{
			
			//Use converter delegate to convert array into another array
			L[] ConvertedArray = Array.ConvertAll(Items, Conversion);
			//Create new list from array (we can do this because the array implements IEnumerable)
			System.Collections.Generic.List<L> ReturnList = new System.Collections.Generic.List<L>(ConvertedArray);
			//All done
			return ReturnList;
			
		}
	}
	
}
