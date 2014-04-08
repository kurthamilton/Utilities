using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Utilities.Reflection
{    
    // It would be nice to make this Generic, so that Compare is Compare(T, T) instead of Compare(object, object).
    internal class Property
    {
        public Type Type { get; private set; }
        public string Name { get; protected set; }
        protected string ChildName { get; private set; }

        private PropertyInfo _propertyInfo;
        protected PropertyInfo PropertyInfo { get { if (_propertyInfo == null) _propertyInfo = GetPropertyInfo(Type, Name); return _propertyInfo; } }

        private Property _child;
        public virtual Property Child 
        { 
            get 
            {
                if (_child == null && ChildName != null)
                    _child = new Property(PropertyInfo.PropertyType, ChildName);
                return _child;
            } 
        }

        public Property(Type type, string name)
        {
            Type = type;
            Name = name;

            int propertyDelimiter = Name.IndexOf(".");
            if (propertyDelimiter > -1)
            {
                ChildName = name.Substring(propertyDelimiter + 1);
                Name = name.Substring(0, propertyDelimiter);                    
            }         
        }
        
        public object GetValue(object obj)
        {            
            if (!string.IsNullOrEmpty(Name) && obj != null)
            {
                object value = PropertyInfo.GetValue(obj, null);
                if (Child != null)
                    value = Child.GetValue(value);
                return value;
            }
            return null;
        }

        public virtual int Compare(object x, object y)
        {
            if (x == null || y == null)
            {                
                if (y != null)
                    return -1;
                else if (x != null)
                    return 1;
                return 0;
            }

            // wouldn't need to do this check with Generics, but not sure how to set up dynamic generics
            if (x.GetType() != Type || y.GetType() != Type)
                throw new Exception("Incompatible Types");
            
            object xValue = GetValue(x);
            object yValue = GetValue(y);

            if (!object.Equals(xValue, yValue))
            {
                if (xValue as IComparable != null)
                    return (xValue as IComparable).CompareTo(yValue);
                else if (yValue as IComparable != null)
                    return -1 * (yValue as IComparable).CompareTo(xValue);
                else
                    return string.Compare(xValue.ToString(), yValue.ToString());
            }                
            return 0;
        }

        /***********************************
        * Static methods
        ************************************/

        public static IEnumerable<Property> GetTypePropertiesWithAttribute(Type type, Type attributeType)
        {
            List<Property> propertiesWithAttribute = new List<Property>();

            IEnumerable<PropertyInfo> properties = type.GetProperties();
            foreach (var property in properties)
            {
                if (property.GetCustomAttributes(attributeType, true).Count() > 0)
                    propertiesWithAttribute.Add(new Property(type, property.Name));
            }

            return propertiesWithAttribute;
        }

        private static PropertyInfo GetPropertyInfo(Type type, string propertyName)
        {            
            return type.GetProperty(propertyName, BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.Instance);
        }
    }
}
