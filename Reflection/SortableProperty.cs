using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Utilities.Reflection
{
    enum SortDirection : int
    {
        Descending = -1,
        None = 0,
        Ascending = 1
    }

    internal class SortableProperty : Property
    {
        private SortDirection SortDirection { get; set; }

        private Property _child;
        public override Property Child
        {
            get
            {
                if (_child == null && ChildName != null)
                    _child = new SortableProperty(PropertyInfo.PropertyType, ChildName);
                return _child;
            }
        }

        public SortableProperty(Type type, string name)
            : base(type, name)
        {
            int directionDelimeter = name.IndexOf(" ");
            if (directionDelimeter > -1)
            {
                // base class will remove sort direction along with child properties, so use original name to get sort direction                
                string sortDirectionString = name.Substring(directionDelimeter + 1);                
                SortDirection = (sortDirectionString.ToLower() == "desc" ? SortDirection.Descending : SortDirection.Ascending);

                // update base Name if direction hasn't already been removed along with child property
                if (Name.Contains(" "))
                    Name = Name.Substring(0, Name.IndexOf(" "));
            }
            else
            {
                SortDirection = SortDirection.Ascending;
            }
        }

        public override int Compare(object x, object y)
        {            
            int compare = base.Compare(x, y);
            return (int)SortDirection * compare;
        }
    }
}
