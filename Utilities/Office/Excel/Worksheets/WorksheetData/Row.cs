using System;
using System.Collections.Generic;
using System.Linq;

using System.Xml;
using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{
    public class Row : BaseExcel
    {
        public Worksheet Worksheet { get; private set; }
        
        private int _index;
        public int Index { get { return _index; } internal set { _index = value; } }        

        private double _height;
        public double Height { get { if (_height > 0) return _height; return Worksheet.Format.DefaultRowHeight; } private set { _height = value; } }

        private CellCollection _cells;
        public CellCollection Cells { get { if (_cells == null) _cells = new CellCollection(this); return _cells; } private set { _cells = value; } }

        internal int StyleIndex { get; private set; }

        /***********************************
         * CONSTRUCTORS
         ************************************/

        private Row(Worksheet worksheet)
        {
            Worksheet = worksheet;
        }

        internal Row(Worksheet worksheet, int index, int styleIndex = CellFormat.DefaultStyleIndex, double height = -1, bool fullLoad = true)
            : this(worksheet)
        {
            RowCollection.AssertValidIndex(index);

            StyleIndex = styleIndex;            
            Index = index;            
            Height = height;
        }


        /***********************************
         * PUBLIC METHODS
         ************************************/

        /// <summary>
        /// Creates a copy of the Row and assigns to the given Worksheet.
        /// </summary>
        public Row Clone(Worksheet worksheet)
        {
            Row newRow = new Row(worksheet, Index, StyleIndex, Height);
            
            foreach (Cell cell in Cells)
            {
                Cell newCell = cell.Clone(newRow);
                newRow.Cells[newCell.Column.Index] = newCell;
            }

            return newRow;
        }

        /// <summary>
        /// Creates a new copy of the Row.
        /// </summary>
        public Row Clone()
        {
            return Clone(Worksheet);
        }

        /// <summary>
        /// Copies and inserts the Row to the given row index.
        /// </summary>
        public Row CopyTo(int insertIndex)
        {
            RowCollection.AssertValidIndex(insertIndex);
            
            Row newRow = Clone();
            newRow.Index = insertIndex;
            Worksheet.Rows.Insert(newRow);
            return newRow;
        }

        /// <summary>
        /// Deletes the Row from the Worksheet.
        /// </summary>
        public void Delete()
        {
            Worksheet.Rows.Delete(Index);
        }


        /***********************************
         * DAL METHODS
         ************************************/


        // Read
        internal static Row ReadRowFromReader(CustomOpenXmlReader reader, Worksheet worksheet)
        {
            Row row = new Row(worksheet);           

            foreach (CustomOpenXmlAttribute attribute in reader.Attributes)
            {
                switch (attribute.LocalName)
                {
                    case "r":
                        row.Index = attribute.GetIntValue();
                        break;
                    case "s":
                        row.StyleIndex = attribute.GetIntValue();
                        break;
                    case "ht":
                        row.Height = attribute.GetDoubleValue();
                        break;
                }
            }            

            row.Cells = CellCollection.ReadCellsFromReader(reader, row);
            return row;
        }

        // Write
        internal static void WriteRowToWriter(CustomOpenXmlWriter<OpenXmlPackaging.WorksheetPart> writer, Row row)
        {
            if (row.Cells.FirstOrDefault(c => c.IsUsed) == null)
                return;

            writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.Row());

            writer.WriteAttribute("r", row.Index);
            if (row.StyleIndex > CellFormat.DefaultStyleIndex) writer.WriteAttribute("s", row.StyleIndex);
            if (row.Height > 0 && row.Height != row.Worksheet.Format.DefaultRowHeight) writer.WriteAttribute("ht", row.Height);

            row.Cells.Action(c => Cell.WriteCellToWriter(writer, c));
            
            writer.WriteEndElement();   // Row
        }
        
    }
}
