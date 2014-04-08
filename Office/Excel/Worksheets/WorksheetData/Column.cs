using System;
using System.Collections.Generic;
using System.Linq;

namespace Utilities.Office.Excel
{
    public class Column : BaseExcel
    {
        public Worksheet Worksheet { get; private set; }

        private int _index;
        public int Index { get { return _index; } internal set { _index = value; _name = null; } }

        private string _name;
        public string Name { get { if (_name == null) _name = ExcelUtilities.GetColumnIndexConverter().GetColumnNameFromColumnIndex(Index); return _name; } }

        // The worksheet XML structure is effectively <worksheet><row><cell column="column"></row></worksheet>
        // Only reference column's cells by getting them from each row rather than maintaining 2 sets of cell collections referencing the same objects
        public CellCollection Cells { get { return new CellCollection(this); } }
        
        //internal CellFormat ColumnFormat { get; private set; }
        internal int StyleIndex { get; private set; }

        /// <summary>
        /// Note that the rendered column width is based on the font within that column. This value does not take font into account, so is only approximate.
        /// </summary>
        public double Width 
        { 
            get 
            {
                if (Worksheet.MaximumColumnWidth == 0 || _width <= Worksheet.MaximumColumnWidth) 
                    return _width; 
                return Worksheet.MaximumColumnWidth; 
            } 
            set 
            { 
                // through testing this seems to be out by a constant value, so just add a constant to get nearer to the expected value.
                if (value > 0)
                    _width = value + 0.83; 
            } 
        }
        private double _width;

        internal const double DefaultColumnWidth = 9.140625;
        internal bool IsBestFit { get; private set; }
        internal bool IsCustomWidth { get; private set; }


        /***********************************
         * CONSTRUCTORS
         ************************************/


        public Column(Worksheet worksheet, int index)
            : this (worksheet, index, CellFormat.DefaultStyleIndex, DefaultColumnWidth, false, false)
        {
        }

        // internal for time being - until full styling is required
        internal Column(Worksheet worksheet, int index, int columnFormatId, double width, bool isBestFit, bool isCustomWidth)
        {
            Worksheet = worksheet;
            StyleIndex = columnFormatId;

            ColumnCollection.AssertValidIndex(index);
            Index = index;
            _width = width; // Publicly setting the Width accounts for measuring differences, so bypass by directly setting here.
            IsBestFit = isBestFit;
            IsCustomWidth = isCustomWidth;
        }

        internal Column (ColumnRange columnRange, int index)
            : this (columnRange.Worksheet, index, columnRange.StyleIndex, columnRange.Width, columnRange.IsBestFit, columnRange.IsCustomWidth)
        {
            if (index < columnRange.MinIndex || index > columnRange.MaxIndex)
                throw new ArgumentOutOfRangeException();
        }

        /***********************************
         * PUBLIC METHODS
         ************************************/


        /// <summary>
        /// Sets the Width of the Column based on the cell contents.
        /// </summary>
        public void AutoFit()
        {
            double maxRequiredWidth = Cells.Max(c => c.RequiredWidth);

            if (maxRequiredWidth > 0 && maxRequiredWidth != DefaultColumnWidth)
            {
                Width = maxRequiredWidth;

                IsCustomWidth = true;
                IsBestFit = true;
            }            
        }

        /// <summary>
        /// Creates a copy of the Column and assigns to the given Worksheet.
        /// </summary>
        public Column Clone(Worksheet worksheet)
        {
            Column newColumn = new Column(worksheet, Index, StyleIndex, Width, IsBestFit, IsCustomWidth);            
            return newColumn;
        }

        /// <summary>
        /// Creates a new copy of the Column.
        /// </summary>
        public Column Clone()
        {
            return Clone(Worksheet);
        }

        /// <summary>
        /// Copies and inserts the Column to the given column index.
        /// </summary>
        public Column CopyTo(int insertIndex)
        {
            if (insertIndex > 0 && insertIndex < ColumnCollection.MaxValue)
            {
                Column newColumn = Clone(Worksheet);
                newColumn.Index = insertIndex;
                Worksheet.Columns.Insert(newColumn);
                Cells.Action(c => c.CopyTo(Worksheet.Rows[c.RowIndex], newColumn));
                return newColumn;
            }
            else
                throw new ArgumentOutOfRangeException();
        }

        /// <summary>
        /// Determines whether the specified Column is equal to the current Column in all respects but index and Worksheet.
        /// </summary>
        public bool Similar(Column column)
        {
            return (
                //column.EqualsStyle(this) &&
                column.IsBestFit == IsBestFit &&
                column.IsCustomWidth == IsCustomWidth &&
                column.Width == Width);
        }

        /// <summary>
        /// Deletes the Column from the Worksheet.
        /// </summary>
        public void Delete()
        {
            Worksheet.Columns.Delete(Index);            
        }
    }
}
