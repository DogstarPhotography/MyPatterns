using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Generics
{
    public class Data
    {
        public string Name { get; set; }
        public string Description { get; set; }
    }

    public class MyDictionary : IDictionary<string, Data>
    {

        //Use a generic dictionary as our internal store
        private Dictionary<string, Data> myDictionary = new Dictionary<string, Data>();

        #region IDictionary<string,Data> Members

        void IDictionary<string, Data>.Add(string key, Data value)
        {
            //TODO: Add any custom code here before adding to internal dictionary
            myDictionary.Add(key, value);
        }

        bool IDictionary<string, Data>.ContainsKey(string key)
        {
            return myDictionary.ContainsKey(key);
        }

        ICollection<string> IDictionary<string, Data>.Keys
        {
            get { throw new NotImplementedException(); }
        }

        bool IDictionary<string, Data>.Remove(string key)
        {
            myDictionary.Remove(key);
            return true;
        }

        bool IDictionary<string, Data>.TryGetValue(string key, out Data value)
        {
            return myDictionary.TryGetValue(key, out value);
        }

        ICollection<Data> IDictionary<string, Data>.Values
        {
            get { return myDictionary.Values; }
        }

        Data IDictionary<string, Data>.this[string key]
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

        #endregion

        #region ICollection<KeyValuePair<string,Data>> Members

        void ICollection<KeyValuePair<string, Data>>.Add(KeyValuePair<string, Data> item)
        {
            myDictionary.Add(item.Key, item.Value);
        }

        void ICollection<KeyValuePair<string, Data>>.Clear()
        {
            myDictionary.Clear();
        }

        bool ICollection<KeyValuePair<string, Data>>.Contains(KeyValuePair<string, Data> item)
        {
            return myDictionary.Contains(item);
        }

        void ICollection<KeyValuePair<string, Data>>.CopyTo(KeyValuePair<string, Data>[] array, int arrayIndex)
        {
            ((ICollection<KeyValuePair<string, Data>>)myDictionary).CopyTo(array, arrayIndex);
        }

        int ICollection<KeyValuePair<string, Data>>.Count
        {
            get { return myDictionary.Count; }
        }

        bool ICollection<KeyValuePair<string, Data>>.IsReadOnly
        {
            get { return false; }
        }

        bool ICollection<KeyValuePair<string, Data>>.Remove(KeyValuePair<string, Data> item)
        {
            myDictionary.Remove(item.Key);
            return true;
        }

        #endregion

        #region IEnumerable<KeyValuePair<string,Data>> Members

        IEnumerator<KeyValuePair<string, Data>> IEnumerable<KeyValuePair<string, Data>>.GetEnumerator()
        {
            return myDictionary.GetEnumerator();
        }

        #endregion

        #region IEnumerable Members

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return myDictionary.GetEnumerator();
        }

        #endregion
    }
    }
