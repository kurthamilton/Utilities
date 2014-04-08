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
    public class CellCollection : BaseExcel, IEnumerable<Cell>, IOfficeCollection<Cell>
    {
        protected SortedDictionary<int, Cell> Cells = new SortedDictionary<int, Cell>();
        public int Count 
        { 
            get 
            {
                return Cells.Count; 
            } 
        }
        private ParentType parentType;
        private Row parentRow;
        private Column parentColumn;
        private Worksheet Worksheet;

        private int NumberFormatId { get; set; }
        internal int CellFormatId { get; set; }
        
        private enum ParentType : int
        {
            Row,
            Column
        }


        /***********************************
         * CONSTRUCTORS
         ************************************/

        internal CellCollection(Row row)
        {
            parentRow = row;
            parentType = ParentType.Row;
            Worksheet = row.Worksheet;
        }
        internal CellCollection(Column column)
        {
            parentColumn = column;
            parentType = ParentType.Column;
            Worksheet = column.Worksheet;
            GetCells(column);
        }


        /***********************************
         * PUBLIC PROPERTIES
         ************************************/

        // alignment
        public VerticalAlignmentOptions VerticalAlignment { get; set; }
        public bool WrapText { get; set; }
        // border
        internal int BorderId { get; set; }
        // fill
        internal int FillId { get; set; }
        // font
        public Font Font { get; private set; }
        // number format


        /***********************************
         * PUBLIC METHODS
         ************************************/


        public Cell this[int index]
        {
            get
            {
                if (!Contains(index))
                    AddCellToCollection(index);
                return Cells[index];
            }

            set
            {
                if (parentType == ParentType.Column)
                    Worksheet.Rows[index].Cells[GetParentIndex()] = value;
                else
                    Cells[index] = value;
            }
        }

        /// <summary>
        /// Clears the cell formats and values for each Cell in the current Cell Collection.
        /// </summary> 
        public void Clear()
        {
            foreach (Cell cell in Cells.Values)
            {
                cell.Clear();
            }
        }

        /// <summary>
        /// Determines whether the given Cell exists within the current Cell Collection.
        /// </summary>
        public bool Contains(Cell cell)
        {
            return (Contains(GetCellIndex(cell)));
        }

        /// <summary>
        /// Determines whether a Cell exists at the given index within the current Cell Collection.
        /// </summary>
        public bool Contains(int index)
        {
            return (Cells.ContainsKey(index));
        }

        /// <summary>
        /// Insert a new blank Cell at the given index within the current Cell Collection.
        /// <para> </para>Only implemented for Row.Cells.
        /// </summary>
        public Cell Insert(int index)
        {
            if (parentType == ParentType.Row)
            {
                Cell newCell = new Cell(parentRow, Worksheet.Columns[index]);
                return Insert(newCell);
            }
            else
                throw new NotImplementedException();
        }

        /// <summary>
        /// Perform the given action on all Cells.
        /// </summary>
        /// <param name="action"></param>
        public void Action(Action<Cell> action)
        {
            foreach (Cell cell in this)
            {
                action(cell);
            }
        }

        /// <summary>
        /// Deletes the Cell from the current Cell Collection with the given index.
        /// <para> </para>Only implemented for Row.Cells.
        /// </summary>
        public void Delete(int index)
        {
            if (parentType == ParentType.Row)
            {
                List<Cell> affectedCells = (from c in Cells.Values where c.Column.Index >= index select c).ToList();

                foreach (Cell cell in affectedCells)
                {
                    RemoveCellFromCollection(cell.Column.Index);

                    if (cell.Column.Index != index)
                    {
                        cell.Column = Worksheet.Columns[cell.Column.Index - 1];
                        AddCellToCollection(cell.Column.Index, cell);
                    }
                }
            }
            else
                throw new NotImplementedException();   
        }

        /// <summary>
        /// Returns the first Cell in the Cell Collection with the given cell value.
        /// </summary>
        public Cell FindCellByValue(object value, bool ignoreCase)
        {
            List<Cell> matches = FindCellsByValue(value, ignoreCase, true);
            if ((matches) != null)
                if (matches.Count > 0)
                    return matches[0];

            return null;
        }

        /// <summary>
        /// Returns all Cells in the Cell Collection with the given cell value.
        /// </summary>
        public List<Cell> FindCellsByValue(object value, bool ignoreCase)
        {
            return FindCellsByValue(value, ignoreCase, false);
        }        

        /// <summary>
        /// Gets the first Cell in the current Cell collection that contains a Cell with a value. Returns null if there are no cell values.
        /// </summary>
        public Cell GetFirstNonBlankCell()
        {
            foreach (Cell cell in Cells.Values)
            {
                if (cell.Text != string.Empty)
                    return cell;
            }

            return null;
        }

        // Implement IEnumerable
        public IEnumerator<Cell> GetEnumerator()
        {
            PadBlankCells();
            return new GenericEnumerator<Cell>(Cells);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        /***********************************
         * INTERNAL METHODS
         ************************************/

        internal Cell Insert(Cell newCell)
        {
            if (parentType == ParentType.Row)
            {
                List<KeyValuePair<int, Cell>> affectedCellIndexes = Cells.Where(c => c.Value.ColumnIndex >= newCell.ColumnIndex).ToList();
                affectedCellIndexes.Reverse();

                foreach (KeyValuePair<int, Cell> affectedCellIndex in affectedCellIndexes)
                {
                    Cell cell = affectedCellIndex.Value;
                    int index = affectedCellIndex.Key;

                    if (index != cell.ColumnIndex)
                    {
                        RemoveCellFromCollection(index);
                        cell.Column = Worksheet.Columns[index + 1];
                        AddCellToCollection(cell.ColumnIndex, cell);
                    }
                }

                AddCellToCollection(newCell.ColumnIndex, newCell);
                return Cells[newCell.ColumnIndex];
            }
            else
                throw new NotImplementedException();
        }

        internal List<Cell> FindCellsByValue(object value, bool ignoreCase, bool firstMatchOnly)
        {
            List<Cell> matches = new List<Cell>();

            bool isStringSearch = (value.GetType() == typeof(string));

            foreach (Cell cell in Cells.Values)
            {
                if (!ignoreCase && isStringSearch)
                {
                    if (string.Compare(value.ToString(), cell.Text, ignoreCase) == 0)
                    {
                        matches.Add(cell);
                        if (firstMatchOnly)
                            break;
                    }
                }
                else
                {
                    if (cell.Value.Equals(value))
                    {
                        matches.Add(cell);
                        if (firstMatchOnly)
                            break;
                    }
                }
            }

            if (matches.Count > 0)
                return matches;
            else
                return null;
        }

        internal void AddToCollection(Cell cell)
        {
            if (!Contains(cell))
            {
                if (parentType == ParentType.Row)
                    AddCellToCollection(cell.Column.Index, cell);
                else if (parentType == ParentType.Column)
                    AddCellToCollection(cell.Row.Index, cell);
                else
                    throw new ArgumentOutOfRangeException();
            }
        }

        /// <summary>
        /// Removes a Cell from the current Cell Collection for the given index.
        /// </summary>
        internal void Remove(int index)
        {
            RemoveCellFromCollection(index);
        }


        /***********************************
         * PRIVATE METHODS
         ************************************/        

        private void GetCells(Column column)
        {
            foreach (Row row in column.Worksheet.Rows)
            {
                AddCellToCollection(row.Index, row.Cells[column.Index]);
            }
        }

        private void PadBlankCells()
        {
            int indexFrom;
            int indexTo;

            switch (parentType)
            {
                case ParentType.Row:
                    indexFrom = ColumnCollection.MinValue;
                    indexTo = Worksheet.Columns.Count;
                    break;
                case ParentType.Column:
                    indexFrom = RowCollection.MinValue;
                    indexTo = Worksheet.Rows.Count;
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }

            for (int index = indexFrom; index <= indexTo; index++)
            {
                AddCellToCollection(index);
            }
        }

        private void AddCellToCollection(int index)
        {
            if (!Contains(index))
            {
                Cell newCell;

                switch (parentType)
                {
                    case ParentType.Row:
                        Column column = parentRow.Worksheet.Columns[index];
                        newCell = new Cell(parentRow, column, "", parentRow.StyleIndex);
                        break;
                    case ParentType.Column:
                        Row row = parentColumn.Worksheet.Rows[index];
                        newCell = new Cell(row, parentColumn, "", row.StyleIndex);
                        break;
                    default:
                        throw new ArgumentOutOfRangeException();
                }

                AddCellToCollection(index, newCell);
            }
        }

        private void AddCellToCollection(int index, Cell cell)
        {
            if (!Contains(index))
                Cells.Add(index, cell);
            else
                Cells[index] = cell;
        }

        private void RemoveCellFromCollection(int index)
        {
            if (Cells.ContainsKey(index))
                Cells.Remove(index);
        }

        private int GetCellIndex(Cell cell)
        {
            switch (parentType)
            {
                case ParentType.Row:
                    return cell.Column.Index;
                case ParentType.Column:
                    return cell.Row.Index;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        private int GetParentIndex()
        {
            switch (parentType)
            {
                case ParentType.Row:
                    return parentRow.Index;
                case ParentType.Column:
                    return parentColumn.Index;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }



        /***********************************
         * DAL METHODS
         ************************************/


        // Read
        internal static CellCollection ReadCellsFromReader(CustomOpenXmlReader reader, Row row)
        {
            CellCollection cells = new CellCollection(row);

            while (reader.ReadToEndElement<OpenXmlSpreadsheet.Row>())
            {
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.Cell>())
                {
                    cells.AddToCollection(Cell.ReadCellFromReader(reader, row));
                }
            }

            return cells;
        }


        // Write
        
    }
}