using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using System.Xml;
using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{
    public class ColumnCollection : BaseExcel, IEnumerable<Column>, IOfficeCollection<Column>
    {
        public const int MinValue = 1;
        public const int MaxValue = 16384;

        private SortedDictionary<int, Column> _columns;
        private SortedDictionary<int, Column> Columns { get { if (_columns == null) _columns = new SortedDictionary<int,Column>(); return _columns; } set { _columns = value; } }
        private SortedDictionary<int, ColumnRange> _columnRanges;
        private SortedDictionary<int, ColumnRange> ColumnRanges { get { if (_columnRanges == null) _columnRanges = new SortedDictionary<int,ColumnRange>(); return _columnRanges; } set { _columnRanges = value; } }

        private Worksheet Worksheet { get; set; }
        
        public int Count 
        { 
            get 
            {                
                return Columns.Count; 
            } 
        }

        /***********************************
         * CONSTRUCTORS
         ************************************/

        internal ColumnCollection(Worksheet worksheet)
        {
            Worksheet = worksheet;
        }

        private ColumnCollection(Worksheet worksheet, ColumnCollection columnCollection)
        {
            Worksheet = worksheet;

            foreach (ColumnRange columnRange in columnCollection.ColumnRanges.Values)
            {
                AddColumnRangeToCollection(columnRange.Clone(worksheet));
            }

            foreach (Column column in columnCollection)
            {
                AddColumnToCollection(column.Clone(worksheet));
            }
        } 


        /***********************************
         * PUBLIC METHODS
         ************************************/     


        public Column this[int index]
        {
            get
            {
                if (!Contains(index))
                    AddColumnToCollection(Worksheet, index);
                return Columns[index];
            }
        }


        /// <summary>
        /// Sets the Width of each Column based on the cell contents.
        /// </summary>
        public void AutoFit()
        {
            foreach (Column column in this)
            {
                column.AutoFit();
            }
        }


        /// <summary>
        /// Determines whether the given Column index exists within the Worksheet Columns.
        /// </summary>
        public bool Contains(int index)
        {
            return Columns.ContainsKey(index);            
        }

        /// <summary>
        /// Inserts a new empty Column at the given index. 
        /// <para></para>Exceptions:<para></para>ArgumentOutOfRangeException
        /// </summary>
        public Column Insert(int index)
        {
            Column newColumn = new Column(Worksheet, index);
            return Insert(newColumn);
        }

        internal Column Insert(Column newColumn)
        {
            int index = newColumn.Index;

            // Increment column indexes >= insert index to vacate the column for the given index

            List<Column> affectedColumns = Columns.Values.Where(c => c.Index >= index).ToList();

            if (affectedColumns.Count > 0)
            {
                // update columns in reverse order to avoid creating conflicting column index keys when updating column collection
                affectedColumns.Reverse();

                foreach (Column column in affectedColumns)
                {
                    RemoveColumnFromCollection(column);
                    column.Index++;
                    AddColumnToCollection(column);
                }


                foreach (Row row in Worksheet.Rows)
                {
                    row.Cells.Insert(index);
                }
            }

            // Insert new column for given index
            AddColumnToCollection(newColumn);

            // return new column
            return Columns[index];
        }

        /// <summary>
        /// Deletes the Column with the given index from the Worksheet.
        /// </summary>
        public void Delete(int index)
        {
            if (Contains(index))
            {
                Column deletedColumn = Columns[index];

                // delete columns
                List<Column> affectedColumns = (from c in Columns.Values where c.Index >= index select c).ToList();

                foreach (Column column in affectedColumns)
                {
                    CellCollection columnCells = column.Cells;
                    columnCells.Action(c => c.RemoveFromRowCollection());                    
                    RemoveColumnFromCollection(column);

                    if (column.Index != index)
                    {
                        column.Index--;
                        AddColumnToCollection(column);

                        columnCells.Action(c => c.AddToRowCollection());
                    }
                }
            }
        }

        // Implement IEnumerable
        public IEnumerator<Column> GetEnumerator()
        {
            return new GenericEnumerator<Column>(Columns);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }


        /***********************************
         * INTERNAL METHODS
         ************************************/


        internal ColumnCollection Clone(Worksheet worksheet)
        {
            ColumnCollection newColumnCollection = new ColumnCollection(worksheet, this);
            return newColumnCollection;
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

        private void LoadColumnRanges()
        {
            _columnRanges = new SortedDictionary<int, ColumnRange>();

            List<ColumnRange> columnRanges = ColumnRange.GetColumnRangesFromWorksheet(Worksheet);
            foreach (ColumnRange columnRange in columnRanges)
            {
                AddColumnRangeToCollection(columnRange);
            }
        }

        private void LoadColumns()
        {
            _columns = new SortedDictionary<int, Column>();

            // get column ranges before last possible column
            foreach (ColumnRange columnRange in ColumnRanges.Values)
            {
                if (columnRange.MaxIndex < ColumnCollection.MaxValue)
                {
                    for (int index = columnRange.MinIndex; index <= columnRange.MaxIndex; index++)
                    {
                        Column column = new Column(columnRange, index);
                        column.Width = columnRange.Width;
                        AddColumnToCollection(column);
                        if (index >= ColumnCollection.MaxValue)
                            break;
                    }
                }
                else 
                    break;
            }
        }

        private void AddColumnToCollection(Worksheet worksheet, int index)
        {
            Column column = new Column(worksheet, index);
            AddColumnToCollection(column);
        }

        private void AddColumnToCollection(Column column)
        {
            if (!Contains(column.Index))
                Columns.Add(column.Index, column);
            else
                Columns[column.Index] = column;
        }

        private void AddColumnRangeToCollection(ColumnRange columnRange)
        {
            if (!ColumnRanges.ContainsKey(columnRange.MinIndex))
                ColumnRanges.Add(columnRange.MinIndex, columnRange);
            else
                ColumnRanges[columnRange.MinIndex] = columnRange;
        }

        private void RemoveColumnFromCollection(Column column)
        {
            if (Contains(column.Index))
            {                
                Columns.Remove(column.Index);
            }
        }
        

        /***********************************
        * DAL METHODS
        ************************************/


        // Read
        internal static ColumnCollection ReadColumnsFromReader(CustomOpenXmlReader reader, Worksheet worksheet)
        {
            ColumnCollection columns = new ColumnCollection(worksheet);

            while (reader.ReadToEndElement<OpenXmlSpreadsheet.Columns>())
            {
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.Column>())
                {
                    ColumnRange columnRange = ColumnRange.ReadColumnRangeFromReader(reader, worksheet);
                    columns.AddColumnRangeToCollection(columnRange);
                }
            }

            columns.LoadColumns();

            return columns;
        }

        // Write

        internal static void WriteColumnsToWriter(CustomOpenXmlWriter<OpenXmlPackaging.WorksheetPart> writer, ColumnCollection columns)
        {
            // derive new column ranges
            List<ColumnRange> columnRanges = ColumnRange.GetColumnRangesFromWorksheet(columns.Worksheet);            

            if (columnRanges.Count > 0)
            {
                // add last column range
                ColumnRange lastColumnRange = columnRanges.Last();
                if (lastColumnRange.MaxIndex < ColumnCollection.MaxValue && columns.ColumnRanges.Count > 0)
                {
                    columnRanges.Add(new ColumnRange(columns.Worksheet, lastColumnRange.MaxIndex + 1, ColumnCollection.MaxValue, lastColumnRange.StyleIndex, Column.DefaultColumnWidth, false, false));
                    //columnRanges.Last().MinIndex = lastColumnRange.MaxIndex + 1;
                    //columnRanges.Last().MaxIndex = ColumnCollection.MaxValue;
                }

                writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.Columns());

                foreach (ColumnRange columnRange in columnRanges)
                {
                    ColumnRange.WriteColumnRangeToWriter(writer, columnRange);
                }

                writer.WriteEndElement();   // Columns
            }
        }
        


    }
}
