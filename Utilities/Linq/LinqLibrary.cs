using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Utilities.Reflection;

namespace Utilities.Linq
{
    public static class LinqLibrary
    {
        /// <summary>
        /// Sorts the elements of a sequence in order according to a sort string.
        /// </summary>        
        /// <param name="sortString">The comma separated sort key. Append each comma separated value with " asc" or " desc" to specify order.</param>
        public static IEnumerable<T> OrderBy<T>(this IEnumerable<T> items, string sortString)
        {
            if (items.FirstOrDefault() != null && !string.IsNullOrEmpty(sortString))
            {                
                string[] sortProperties = sortString.Split(',');

                List<T> list = items.ToList();
                list.Sort(new PropertyComparer<T>(sortProperties));
                return list;
            }
            return items;
        }

        /// <summary>
        /// Gets the index of the first item that satisfies the given condition. Returns -1 if the condition is not satisfied.
        /// </summary>        
        public static int IndexOf<T>(this IEnumerable<T> items, Func<T, bool> predicate)
        {
            int index = 0;
            foreach (var item in items)
            {                
                if (predicate(item)) 
                    return index;
                index++;
            }
            return -1;
        }

        /// <summary>
        /// Returns the element offset from the first element that satisfies the given condition.
        /// </summary>
        public static T Offset<T>(this IEnumerable<T> items, Func<T, bool> predicate, int offset)
        {
            int index = items.IndexOf(predicate);
            int offsetIndex = index + offset;
            if (offsetIndex >= 0 && offsetIndex < items.Count())
                return items.ElementAt(offsetIndex);
            return default(T);
        }

        /// <summary>
        /// Returns the element after the first element that satisfies the given condition.
        /// </summary>
        public static T Next<T>(this IEnumerable<T> items, Func<T, bool> predicate)
        {
            return Offset(items, predicate, 1);
        }

        /// <summary>
        /// Returns the element before the first element that satisfies the given condition.
        /// </summary>
        public static T Previous<T>(this IEnumerable<T> items, Func<T, bool> predicate)
        {
            return Offset(items, predicate, -1);
        }
    }
}
