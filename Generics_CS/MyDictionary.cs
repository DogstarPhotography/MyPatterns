////INSTANT C# NOTE: Formerly VB project-level imports:
//using System;
//using System.Collections;
//using System.Collections.Generic;
//using System.Data;
//using System.Diagnostics;

//namespace Generics
//{
//    public class Data
//    {
//        public string Name {get; set;}
//        public string Description {get; set;}
//    }
///// <summary>
///// Template for implementing IDictionary(Of TKey, TValue).
///// </summary>
///// <remarks></remarks>
//    public class MyDictionary : IDictionary<string, Data>
//    {
//#region Implementation
//        //Use a generic dictionary as our internal store
//        private Dictionary<string, Data> myDictionary = new Dictionary<string, Data>();

//#region Implements IDictionary<string, Data>

//        void System.Collections.Generic.IDictionary<string, Data>.Add(string key, Data value)
//        {
//            this.Add1(key, value);
//        }
//        public void Add1(string key, Data value)
//        {
//            //TODO: Add any custom code here before adding to internal dictionary
//            myDictionary.Add(key, value);
//        }

//        public bool ContainsKey(string key)
//        {
//            return myDictionary.ContainsKey(key);
//        }

//        public Data this[string key]
//        {
//            get
//            {
//                //TODO: Add any custom code here before reading from internal dictionary
//                return myDictionary[key];
//            }
//            set
//            {
//                //TODO: Add any custom code here before adding to internal dictionary
//                myDictionary[key] = value;
//            }
//        }

//        bool System.Collections.Generic.IDictionary<string, Data>.Remove(string key)
//        {
//            return this.Remove1(key);
//        }
//        public bool Remove1(string key)
//        {
//            myDictionary.Remove(key);
////INSTANT C# NOTE: Inserted the following 'return' since all code paths must return a value in C#:
//            return false;
//        }

//        public bool TryGetValue(string key, ref Data value)
//        {
//            return myDictionary.TryGetValue(key, out value);
//        }

//        public System.Collections.Generic.ICollection<Data> Values
//        {
//            get
//            {
//                return myDictionary.Values;
//            }
//        }

//#endregion

//#region Implements ICollection<string, Data>

////INSTANT C# TODO TASK: The following line could not be converted:
//        Public Sub Add(item As System.Collections.Generic.KeyValuePair<string, Data>) Implements System.Collections.Generic.ICollection<System.Collections.Generic.KeyValuePair<string, Data>>.Add
//            //TODO: Add any custom code here before adding to internal dictionary
//            myDictionary.Add(item.Key, item.Value);
//        }

////INSTANT C# TODO TASK: The following line could not be converted:
//        Public Sub Clear() Implements System.Collections.Generic.ICollection<System.Collections.Generic.KeyValuePair<string, Data>>.Clear
//            myDictionary.Clear();
//        }

//        public bool Contains(System.Collections.Generic.KeyValuePair<string, Data> item)
//        {
//            return myDictionary.ContainsKey(item.Key);
//        }

//        public void CopyTo(System.Collections.Generic.KeyValuePair<string, Data>[] array, int arrayIndex)
//        {
//            ((ICollection<KeyValuePair<string, Data>>)myDictionary).CopyTo(array, arrayIndex);
//        }

//        public int Count
//        {
//            get
//            {
//                return myDictionary.Count;
//            }
//        }

//        public bool IsReadOnly
//        {
//            get
//            {
//                return false;
//            }
//        }

//        public bool Remove(System.Collections.Generic.KeyValuePair<string, Data> item)
//        {
//            myDictionary.Remove(item.Key);
////INSTANT C# NOTE: Inserted the following 'return' since all code paths must return a value in C#:
//            return false;
//        }

//        public System.Collections.Generic.ICollection<string> Keys
//        {
//            get
//            {
//                return myDictionary.Keys;
//            }
//        }

//#endregion

//#region Implements IEnumerable<string, Data>, IEnumerable

//        public System.Collections.Generic.IEnumerator<System.Collections.Generic.KeyValuePair<string, Data>> GetEnumerator()
//        {
//            return myDictionary.GetEnumerator();
//        }

//        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
//        {
//            return this.GetEnumerator1();
//        }
//        public System.Collections.IEnumerator GetEnumerator1()
//        {
//            return myDictionary.GetEnumerator();
//        }

//#endregion

//#endregion

//    }
//}