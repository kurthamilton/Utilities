using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;

namespace Utilities
{
    public static class Helpers
    {
        public static T GetEnumValueFromDescription<T>(string description) where T : struct
        {
            Type t = typeof(T);
            Array values = Enum.GetValues(t);
            foreach (var value in values)
            {
                if (string.Compare(Enum.GetName(t, value), description, false) == 0)
                {
                    T outputValue;
                    Enum.TryParse<T>(value.ToString(), out outputValue);
                    return outputValue;
                }
            }
            return new T();
        }

        public static string ToCamelCase(string value)
        {
            if (!string.IsNullOrEmpty(value))
            {
                value = string.Concat(value.Substring(0, 1).ToLower(), value.Substring(1));
            }
            return value;
        }

        public static HttpPostedFileBase GetHttpPostedFileBaseFromHttpPostedFile(HttpPostedFile postedFile)
        {
            return new HttpPostedFileWrapper(postedFile);
        }
    }
}
