using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace Utilities.Office.Excel
{
    public class ColumnIndexConverter : BaseExcel
    {        
        protected Dictionary<int, string> indexedColumnsByNumber = new Dictionary<int, string>();
        protected Dictionary<string, int> indexedColumnsByName = new Dictionary<string, int>();

        public ColumnIndexConverter()
        {
        }

        public int GetColumnIndexFromCellAddress(string cellAddress)
        {            
            // Leave just column name
            string columnName = GetColumnNameFromCellAddress(cellAddress);

            if (columnName != "")
            {
                if (!indexedColumnsByName.ContainsKey(columnName))
                {
                    int columnIndex = GetColumnIndexFromColumnName(columnName);
                    UpdateIndexes(columnIndex, columnName);
                }

                return indexedColumnsByName[columnName];
            }
            return 0;
        }

        public string GetColumnNameFromColumnIndex(int columnIndex)
        {
            if (columnIndex > 0)
            {
                if (!indexedColumnsByNumber.ContainsKey(columnIndex))
                {
                    string columnName = GetColumnNameFromColumnIndex1(columnIndex);
                    UpdateIndexes(columnIndex, columnName);
                }

                return indexedColumnsByNumber[columnIndex];
            }
            return "";
        }

        public void GetCellIndexesFromCellAddress(string cellAddress, out int rowIndex, out int columnIndex)
        {
            rowIndex = 0;
            columnIndex = GetColumnIndexFromCellAddress(cellAddress);
            if (columnIndex > 0)
            {
                string columnName = GetColumnNameFromColumnIndex(columnIndex);
                int.TryParse(cellAddress.Replace(columnName, ""), out rowIndex);
            }
        }

        protected string GetColumnNameFromCellAddress(string cellAddress)
        {
            // don't handle ranges
            if (cellAddress.Contains(":"))
                return "";

            return Regex.Replace(cellAddress.ToUpper(), @"\d", "");
        }                

        protected void UpdateIndexes(int columnIndex, string columnName)
        {            
            if (!indexedColumnsByNumber.ContainsKey(columnIndex))
                indexedColumnsByNumber.Add(columnIndex, columnName);
            if (!indexedColumnsByName.ContainsKey(columnName))
                indexedColumnsByName.Add(columnName, columnIndex);
        }

        /************************************************
         STATIC METHODS
         ************************************************/


        protected static int GetBaseAscii()
        {
            return Convert.ToInt32('A') - 1;
        }

        protected static string GetColumnNameFromColumnIndex1(int columnIndex)
        {
            StringBuilder columnName = new StringBuilder("");

            int baseAscii = GetBaseAscii();

            while (columnIndex > 0)
            {
                int mod = (columnIndex) % 26;
                if (mod == 0)
                    mod = 26;

                columnName.Insert(0, Convert.ToChar(baseAscii + mod));
                columnIndex = (int)((columnIndex - mod) / 26);
            }

            return columnName.ToString();
        }

        public static int GetColumnIndexFromColumnName(string columnName)
        {
            int columnIndex = 0;

            // Count backwards from length of column name
            int charIndex = columnName.Length;
            int baseAscii = GetBaseAscii();

            foreach (char c in columnName)
            {
                int currentAscii = Convert.ToInt32(c) - baseAscii;

                columnIndex += (int)Math.Pow(26, charIndex - 1) * (currentAscii);

                charIndex--;
            }

            return columnIndex;
        }
        
    }
}
