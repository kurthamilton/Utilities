using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;


namespace Utilities.Office.Excel
{
    public class RowCollection : BaseExcel, IEnumerable<Row>, IOfficeCollection<Row>
    {
        public const int MinValue = 1;
        public const int MaxValue = 1048576;

        private int _currentMaxIndex;

        private Worksheet Worksheet { get; set; }
        
        protected SortedDictionary<int, Row> rows = new SortedDictionary<int, Row>();        
        
        public int Count
        { 
            get 
            {                 
                return rows.Count; 
            } 
        }


        /***********************************
         * CONSTRUCTORS
         ************************************/


        internal RowCollection(Worksheet worksheet)
        {
            Worksheet = worksheet;
        }

        private RowCollection(Worksheet worksheet, RowCollection rowCollection)
            : this (worksheet)
        {
            _currentMaxIndex = rowCollection._currentMaxIndex;
            foreach (Row row in rowCollection)
            {
                AddRowToCollection(row.Clone(worksheet));
            }
        } 


        // ******************************************************************
        // PUBLIC METHODS
        // ******************************************************************//

        /// <summary>
        /// Gets or sets the Worksheet Row for the given Row index.
        /// </summary>
        public Row this[int index]
        {
            get
            {
                PadBlankRows(index);
                return rows[index];                
            }
            set
            {
                PadBlankRows(index);
                rows[index] = value;                
            }
        }

        /// <summary>
        /// Determines whether the given Row index exists within the used Worksheet Rows.
        /// </summary>
        public bool Contains(int index)
        {
            return (index <= _currentMaxIndex);
        }

        /// <summary>
        /// Inserts a new empty Row at the given index. 
        /// </summary>
        /// <param name="index">The insert index of the new Row</param>
        /// <returns>Newly inserted Row</returns>
        public Row Insert(int index)
        {
            return Insert(index, 1)[0];
        }

        /// <summary>
        /// Inserts (n = numberOfRows) empty Rows at the given index. 
        /// <para></para>Exceptions:<para></para>ArgumentOutOfRangeException
        /// </summary>
        /// <param name="index">The insert index of the new Rows</param>
        /// <returns>List of newly inserted Rows</returns>
        public List<Row> Insert(int index, int numberOfRows)
        {
            List<Row> newRows = new List<Row>();

            if (numberOfRows > 0)
            {
                for (int i = 1; i <= numberOfRows; i++)
                {
                    Row newRow = new Row(Worksheet, index);
                    newRows.Add(Insert(newRow));
                }
            }

            return newRows;
        }

        internal Row Insert(Row newRow)
        {
            int index = newRow.Index;

            // Increment row indexes >= insert index to vacate the row for the given index

            if (index <= _currentMaxIndex)
            {
                List<Row> affectedRows = rows.Values.Where(r => r.Index >= index).ToList();
                // update rows in reverse order to avoid creating conflicting row index keys when updating row collection
                affectedRows.Reverse();

                foreach (Row row in affectedRows)
                {
                    RemoveRowFromCollection(row);
                    row.Index++;
                    AddRowToCollection(row);
                }
            }

            // Insert new row for given index
            AddRowToCollection(newRow);

            if (newRow.Index <= Worksheet.FrozenRow) Worksheet.FrozenRow++;

            // return new row
            return rows[index];
        }

        /// <summary>
        /// Gets the first Row in the current Worksheet that contains a Cell with a value. Returns null if there are no cell values.
        /// </summary>
        public Row GetFirstNonBlankRow()
        {
            foreach (Row row in rows.Values)
            {
                if (row.Cells.GetFirstNonBlankCell() != null)
                    return row;
            }

            return null;
        }

        /// <summary>
        /// Returns all Rows in the Worksheet containing the given cell value.
        /// </summary>
        public List<Row> FindRowsByCellValue(object value, bool ignoreCase)
        {
            List<Row> matches = new List<Row>();

            foreach (Row row in rows.Values)
            {
                if (row.Cells.FindCellByValue(value, ignoreCase) != null)
                    matches.Add(row);
            }

            if (matches.Count > 0)
                return matches;
            else
                return null;
        }

        /// <summary>
        /// Deletes the Row with the given index from the Worksheet.
        /// </summary>
        public void Delete(int index)
        {
            if (Contains(index))
            {
                List<Row> affectedRows = (from r in rows.Values where r.Index >= index select r).ToList();

                foreach (Row row in affectedRows)
                {
                    RemoveRowFromCollection(row);

                    if (row.Index != index)
                    {
                        row.Index--;
                        AddRowToCollection(row);
                    }
                }

                _currentMaxIndex--;
                if (Worksheet.FrozenRow > 0) Worksheet.FrozenRow--;
            }
        }

        // Implement IEnumerable
        public IEnumerator<Row> GetEnumerator()
        {
            //PadBlankRows();
            return new GenericEnumerator<Row>(rows);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        /***********************************
         * INTERNAL METHODS
         ************************************/

        internal RowCollection Clone(Worksheet worksheet)
        {
            RowCollection newRowCollection = new RowCollection(worksheet, this);
            return newRowCollection;
        }

        internal static void AssertValidIndex(int index)
        {
            if (!CheckValidIndex(index))
                throw new ArgumentOutOfRangeException();
        }

        internal static bool CheckValidIndex(int index)
        {
            return (index >= MinValue && index <= MaxValue);
        }

        /***********************************
         * PRIVATE METHODS
         ************************************/

        private void PadBlankRows(int indexTo)
        {            
            //int indexFrom = RowCollection.MinValue;
            //if (indexTo == 0)
                //indexTo = _currentMaxIndex;

            if (indexTo > _currentMaxIndex)
            {
                for (int index = _currentMaxIndex + 1; index <= indexTo; index++)
                {
                    //if (!Contains(index))
                        AddRowToCollection(new Row(Worksheet, index));
                }
            }
        }

        private void Add(Row row)
        {
            PadBlankRows(row.Index - 1);
            AddRowToCollection(row);
        }

        private void AddRowToCollection(Row row)
        {
            rows.Add(row.Index, row);
            if (row.Index > _currentMaxIndex)
                _currentMaxIndex = row.Index;
        }

        private void RemoveRowFromCollection(Row row)
        {
            rows.Remove(row.Index);
        }



        /***********************************
         * DAL METHODS
         ************************************/


        // Read
        internal static RowCollection ReadRowsFromReader(CustomOpenXmlReader reader, Worksheet worksheet)
        {
            RowCollection rows = new RowCollection(worksheet);

            while (reader.ReadToEndElement<OpenXmlSpreadsheet.SheetData>())
            {
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.Row>())
                {
                    Row row = Row.ReadRowFromReader(reader, worksheet);
                    rows.Add(row);
                }
            }

            return rows;
        }

        // Write                

        internal static void WriteSheetDimensionToWriter(CustomOpenXmlWriter<OpenXmlPackaging.WorksheetPart> writer, Worksheet worksheet)
        {
            writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.SheetDimension());
            writer.WriteAttribute("ref", string.Format("{0}{1}:{2}{3}", "A", "1", worksheet.Columns.Last().Name, worksheet.Rows.Last().Index));
            writer.WriteEndElement();   // SheetDimension
        }

        internal static void WriteRowsToWriter(CustomOpenXmlWriter<OpenXmlPackaging.WorksheetPart> writer, IEnumerable<Row> rows)
        {
            writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.SheetData());

            foreach (Row row in rows)
            {
                Row.WriteRowToWriter(writer, row);
            }

            writer.WriteEndElement();   // SheetData
        }
    }
}
