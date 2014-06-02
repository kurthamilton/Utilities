using System;

namespace Utilities.Office.Excel
{        
    public static class ExcelUtilities
    {        
        internal static ColumnIndexConverter GetColumnIndexConverter()
        {
            string key = "Excel_ColumnIndexConverter";

            ColumnIndexConverter columnIndexConverter = null;            
            columnIndexConverter = (ColumnIndexConverter)OfficeUtilities.Cache[key];

            if (columnIndexConverter == null)
            {
                columnIndexConverter = new ColumnIndexConverter();
                OfficeUtilities.CacheData(key, columnIndexConverter);
            }
            return columnIndexConverter;
        }

        internal static TimeConverter GetTimeConverter()
        {
            string key = "Excel_TimeConverter";

            TimeConverter timeConverter = null;
            timeConverter = (TimeConverter)OfficeUtilities.Cache[key];

            if (timeConverter == null)
            {
                timeConverter = new TimeConverter();
                OfficeUtilities.CacheData(key, timeConverter);
            }
            return timeConverter;
        }

    }    
}
