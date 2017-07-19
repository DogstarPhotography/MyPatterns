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
	public class Data
	{
		public string Name {get; set;}
		public string Description {get; set;}
	}
	/// <summary>
	/// Template for implementing IDictionary(Of TKey, TValue).
	/// </summary>
	/// <remarks></remarks>
	public class MyDictionary : IDictionary<string, Data>
	{
		
#region Implementation
		//Use a generic dictionary as our internal store
		private Dictionary<string, Data> myDictionary = new Dictionary<string, Data>();
		
#region Implements IDictionary(Of String, Data)
		
		void .Add(string key, Data value)
		{
			this.Add1(key, value);
		}
		
		public void Add1(string key, Data value)
		{
			//TODO: Add any custom code here before adding to internal dictionary
			myDictionary.Add(key, value);
		}
		
		public bool ContainsKey(string key)
		{
			return myDictionary.ContainsKey(key);
		}
		
public Data this[string key]
		{
			get
			{
				//TODO: Add any custom code here before reading from internal dictionary
				return myDictionary[key];
			}
			set
			{
				//TODO: Add any custom code here before adding to internal dictionary
				myDictionary[key] = value;
			}
		}
		
		bool .Remove(string key)
		{
			return this.Remove1(key);
		}
		
		public bool Remove1(string key)
		{
			myDictionary.Remove(key);
		}
		
		public bool TryGetValue(string key, ref Data value)
		{
			return myDictionary.TryGetValue(key, out value);
		}
		
public System.Collections.Generic.ICollection<Data> Values
		{
			get
			{
				return myDictionary.Values;
			}
		}
		
#endregion
		
#region Implements ICollection(Of String, Data)
		
		public void Add(System.Collections.Generic.KeyValuePair<string, Data> item)
		{
			//TODO: Add any custom code here before adding to internal dictionary
			myDictionary.Add(item.Key, item.Value);
		}
		
		public void Clear()
		{
			myDictionary.Clear();
		}
		
		public bool Contains(System.Collections.Generic.KeyValuePair<string, Data> item)
		{
			return myDictionary.ContainsKey(item.Key);
		}
		
		public void CopyTo(System.Collections.Generic.KeyValuePair<string, Data>[] array, int arrayIndex)
		{
			((ICollection<KeyValuePair<string, Data>>) myDictionary).CopyTo(array, arrayIndex);
		}
		
public int Count
		{
			get
			{
				return myDictionary.Count;
			}
		}
		
public bool IsReadOnly
		{
			get
			{
				return false;
			}
		}
		
		public bool Remove(System.Collections.Generic.KeyValuePair<string, Data> item)
		{
			myDictionary.Remove(item.Key);
		}
		
public System.Collections.Generic.ICollection<string> Keys
		{
			get
			{
				return myDictionary.Keys;
			}
		}
		
#endregion
		
#region Implements IEnumerable(Of String, Data), IEnumerable
		
		System.Collections.Generic.IEnumerator<System.Collections.Generic.KeyValuePair<string, Data>> .GetEnumerator()
		{
			return this.GetEnumerator2();
		}
		
		public System.Collections.Generic.IEnumerator<System.Collections.Generic.KeyValuePair<string, Data>> GetEnumerator2()
		{
			return myDictionary.GetEnumerator();
		}
		
		System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
		{
			return this.GetEnumerator1();
		}
		
		public System.Collections.IEnumerator GetEnumerator1()
		{
			return myDictionary.GetEnumerator();
		}
		
#endregion
		
#endregion
		
	}
}
