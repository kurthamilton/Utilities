using System;
using System.Collections;
using System.Collections.Generic;

namespace Utilities.Office.Excel
{
    public class WorksheetCellCollection : BaseExcel, IEnumerable<Cell>
    {
        public Worksheet Worksheet { get; private set; }         

        public WorksheetCellCollection(Worksheet worksheet)
        {
            Worksheet = worksheet;
        }     

        public Cell this[int rowIndex, int columnIndex]
        {
            get
            {
                // Get column and row objects to ensure indexes are valid and padded
                if (!Worksheet.Columns.Contains(columnIndex))
                {
                    Column column = Worksheet.Columns[columnIndex];
                }

                Row row = Worksheet.Rows[rowIndex];
                Cell cell = row.Cells[columnIndex];                
                return cell;
            }

            set
            {
                if (!Worksheet.Columns.Contains(columnIndex))
                {
                    Column column = Worksheet.Columns[columnIndex];
                }
                Row row = Worksheet.Rows[rowIndex];
                
                if (value.RowIndex != rowIndex)
                    value.Row = row;
                if (value.ColumnIndex != columnIndex)
                    value.Column = Worksheet.Columns[columnIndex];

                row.Cells[columnIndex] = value;
            }
        }

        public Cell this[string cellAddress]
        {
            get
            {
                int rowIndex;
                int columnIndex;
                ExcelUtilities.GetColumnIndexConverter().GetCellIndexesFromCellAddress(cellAddress, out rowIndex, out columnIndex);

                if (RowCollection.CheckValidIndex(rowIndex) && ColumnCollection.CheckValidIndex(columnIndex))
                    return this[rowIndex, columnIndex];
                return null;
            }
        }

        /// <summary>
        /// Returns the first Cell in the Cell Collection with the given cell value.
        /// </summary>
        public Cell FindCellByValue(object value, bool ignoreCase)
        {
            foreach (Row row in Worksheet.Rows)
            {
                Cell cell = row.Cells.FindCellByValue(value, ignoreCase);

                if (cell != null)
                    return cell;
            }

            return null;
        }

        /// <summary>
        /// Returns all Cells in the Cell Collection with the given cell value.
        /// </summary>
        public List<Cell> FindCellsByValue(object value, bool ignoreCase)
        {
            List<Cell> matches = new List<Cell>();

            foreach (Row row in Worksheet.Rows)
            {
                List<Cell> rowMatches = row.Cells.FindCellsByValue(value, ignoreCase);
                if (rowMatches != null)
                    matches.AddRange(rowMatches);
            }

            if (matches.Count > 0)
                return matches;
            else
                return null;
        }

        /// <summary>
        /// Clears the cell formats and values for each Cell in the current Worksheet.
        /// </summary>
        public void Clear()
        {
            foreach (Row row in Worksheet.Rows)
            {
                row.Cells.Clear();
            }
        }

        // Implement IEnumerable
        public IEnumerator<Cell> GetEnumerator()
        {
            List<Cell> cells = new List<Cell>();

            foreach (Row row in Worksheet.Rows)
            {
                cells.AddRange(row.Cells);
            }

            return cells.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
