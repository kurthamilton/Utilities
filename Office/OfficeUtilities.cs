using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Caching;

namespace Utilities.Office
{    
    public abstract class OfficeUtilities : BaseOffice
    {        
        internal static Cache Cache
        {
            //get { return HttpContext.Current.Cache; }
            get { return HttpRuntime.Cache; }
        }

        internal static void CacheData(string key, object data)
        {
            if (data != null)
            {
                Cache.Insert(key, data, null, DateTime.Now.AddDays(1), TimeSpan.Zero);
            }
        }                

        internal static int GetFirstUnusedKeyFromCollection<T>(SortedDictionary<int, T> collection)
        {
            int key = 0;
            while (collection.ContainsKey(key))
                key++;
            return key;
        }

        internal static int GetMaxKeyFromCollection<T>(SortedDictionary<int, T> collection)
        {
            if (collection.Count > 0)
                return collection.Keys.Max();
            else 
                return -1;
        }

        internal static List<T> ConvertList<T, B>(IEnumerable<B> businessObjects, Func<B, T> createNew)
        {
            List<T> models = new List<T>();
            if (businessObjects != null)
            {
                foreach (B businessObject in businessObjects)
                    models.Add(createNew(businessObject));
            }
            return models;
        }
    }

    internal class GenericEnumerator<T> : BaseOffice, IEnumerator<T> where T : BaseOffice
    {
        private SortedDictionary<int, T> _values;
        private int _currentIndex;
        private int _maxValue;

        public GenericEnumerator(SortedDictionary<int, T> values)
        {
            _values = values;
            Reset();
        }

        public T Current
        {
            get
            {
                return _values[_currentIndex];
            }
        }

        object IEnumerator.Current
        {
            get { return Current; }
        }

        public void Dispose()
        {
        }

        public bool MoveNext()
        {
            do
            {
                _currentIndex++;

                if (_currentIndex > _maxValue)
                    return false;
            } while (!_values.ContainsKey(_currentIndex));

            return true;
        }

        public void Reset()
        {
            if (_values.Count == 0)
                _currentIndex = 0;
            else
                _currentIndex = _values.Keys.Min() - 1;

            _maxValue = OfficeUtilities.GetMaxKeyFromCollection<T>(_values);
        }
    }


}
