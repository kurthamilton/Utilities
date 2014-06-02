using System;
using System.Collections.Generic;

namespace Utilities.Reflection
{
    internal class PropertyComparer<T> : IComparer<T>
    {
        private IEnumerable<string> PropertyNames { get; set; }

        private List<SortableProperty> _properties;
        public IEnumerable<SortableProperty> Properties
        {
            get
            {
                if (_properties == null)
                {
                    _properties = new List<SortableProperty>();
                    foreach (string propertyName in PropertyNames)
                    {
                        _properties.Add(new SortableProperty(typeof(T), propertyName));
                    }
                }
                return _properties;
            }
        }

        public PropertyComparer(IEnumerable<string> propertyNames)
        {
            PropertyNames = propertyNames;
        }

        public int Compare(T x, T y)
        {
            foreach (var property in Properties)
            {
                int compare = property.Compare(x, y);
                if (compare != 0)
                    return compare;
            }

            return 0;
        }
    }
}
