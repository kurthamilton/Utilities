using System;
using System.Collections.Generic;

using System.Xml;
using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{
    public class Cell : BaseRange
    {
        public Row Row { get; internal set; }
        public int RowIndex { get { return Row.Index; } }
        
        public Column Column { get; internal set; }
        public int ColumnIndex { get { return Column.Index; } }

        public string Address { get { return GetCellAddressByColumnNameByRowIndex(Column.Name, Row.Index); } }
        
        /// <summary>
        /// The raw data value of the cell. Use Cell.Text to get the formatted string value.
        /// </summary>
        public object Value { get { return GetValue(); } set { UpdateCellValue(value); } }
        private object _value;

        private SharedString SharedString { get; set; } // used to store the font values. This is a quick fix. Needs doing properly
        internal object RawValue { get { return _value; } }
        private CellDataType CellDataType { get; set; } 

        /// <summary>
        /// The formatted value of the cell. Use Cell.Value to get the raw data value.
        /// </summary>        
        public string Text { get { return Worksheet.Workbook.Styles.CellFormats[StyleIndex].NumberFormat.GetFormattedValue(GetValue()); } }

        /// <summary>
        /// The number of cells below to merge the current Cell with. Combines with other merge areas below if they already exist. MergeDown currently overrides MergeAcross.
        /// </summary>
        public int MergeDown { get { return _mergeDown; } set { if (value >= 0) _mergeDown = value; } }
        private int _mergeDown;

        /// <summary>
        /// The number of cells to the right to merge the current Cell with. Combines with other merge areas to the right if they already exist. MergeDown currently overrides MergeAcross.
        /// </summary>
        public int MergeAcross { get { return _mergeAcross; } set { if (value >= 0) _mergeAcross = value; } }
        private int _mergeAcross;

        private Cell _hyperlinkToCell;
        public Cell HyperlinkToCell { get { return _hyperlinkToCell; } set { UpdateHyperlink(value); } }

        public bool IsUsed { get { return (Text != "" || StyleIndex > CellFormat.DefaultStyleIndex || MergeDown > 0 || MergeAcross > 0 || HyperlinkToCell != null); } }

        /***********************************
         * CONSTRUCTORS
         ************************************/

        // internal for time being - until full styling is required.
        internal Cell(Row row, Column column, object value, CellDataType cellDataType, int styleIndex)
            : base(row.Worksheet, styleIndex)
        {
            // this constructor should be merged with the one below
            Row = row;
            Column = column;
            CellDataType = cellDataType;
            _value = value;
        }

        internal Cell(Row row, Column column, object value, int styleIndex)
            : base(row.Worksheet, styleIndex)
        {
            Row = row;
            Column = column;
            UpdateCellValue(value);
        }

        public Cell(Row row, Column column, object value)
            : this(row, column, value, CellFormat.DefaultStyleIndex)
        {
            // take row or column format if they exist
            // I can't find anything in the XML files to say which takes precedence if they both exist, so go with row

            //if (row.RowFormat.CellFormatId > CellFormat.DefaultCellFormatId)
            //    CellFormat = row.RowFormat;
            //else if (column.ColumnFormat.CellFormatId > CellFormat.DefaultCellFormatId)
            //    CellFormat = column.ColumnFormat;
            //else
            //    CellFormat = row.Worksheet.Workbook.Styles.CellFormats[CellFormat.DefaultCellFormatId];
        }  

        public Cell(Row row, Column column)
            : this(row, column, string.Empty)
        {                        
        }

        /***********************************
         * PUBLIC PROPERTIES
         ************************************/

        // reset when the cell value or cell format are updated.
        private double _requiredWidth = -1;
        
        /// <summary>
        /// The required width of the Cell based on the Cell contents and Row height.
        /// </summary>
        public double RequiredWidth
        {
            get
            {
                if (_requiredWidth <= 0)
                {
                    double requiredWidth = 0;

                    if (Text != "")
                    {
                        // how many lines can the cell currently accommodate?
                        int lineCount = 0;
                        
                        double characterHeight = Font.GetCharacterHeight(Font.Size);
                        if (characterHeight > 0)
                            lineCount = (int)Math.Floor(Row.Height / characterHeight);
                        if (lineCount <= 0)
                            lineCount = 1;

                        // get full required width for current cell text regardless of row height
                        double characterWidth = Font.GetCharacterWidth(Font.Size);
                        requiredWidth = characterWidth * (double)Text.Length;

                        // split required width based on line count
                        if (lineCount > 1)
                            requiredWidth = requiredWidth / (double)lineCount;                        
                    }

                    _requiredWidth = requiredWidth;
                }
                return _requiredWidth;
            }
        }

        // see note for _requiredWidth
        private double _requiredHeight = -1;

        /// <summary>
        /// The required width of the Cell based on the Cell contents and Row height. Currently assumes text wrapping. 
        /// Use GetRequiredHeight to assume no wrapping.
        /// </summary>
        public double RequiredHeight
        {
            get
            {
                if (_requiredHeight <= 0)
                    _requiredHeight = GetRequiredHeight();
                return _requiredHeight;
            }
        }

        /// <summary>
        /// The required height of the Cell based on the Cell contents and Column width. Will become obsolete when CellFormat.WrapText is implemented.
        /// </summary>
        public double GetRequiredHeight(bool wrapText = true)
        {
            double requiredHeight = 0;

            if (wrapText)
            {
                if (Text != "")
                {
                    double characterWidth = Font.GetCharacterWidth(Font.Size);
                    double requiredWidth = characterWidth * Text.Length;
                    if (Column.Width < requiredWidth)
                    {
                        int lineCount = (int)Math.Ceiling(requiredWidth / Column.Width);
                        double characterHeight = Font.GetCharacterHeight(Font.Size);
                        requiredHeight = characterHeight * (double)lineCount;
                    }
                }
            }

            return requiredHeight;
        }

        /// <summary>
        /// Gets the Cell offset by the given values from the current Cell.
        /// </summary>
        /// <returns></returns>
        public Cell Offset(int rowIndex = 0, int columnIndex = 0)
        {
            return Worksheet.Cells[Row.Index + rowIndex, Column.Index + columnIndex];
        }

        /***********************************
         * PUBLIC METHODS
         ************************************/

        /// <summary>
        /// Add a value with a given font to the current cell value.
        /// </summary>
        public void AppendFontString(Font font, string value)
        {            
            // convert current cell to shared string if it isn't already so
            if (CellDataType != Excel.CellDataType.SharedString)
                UpdateCellValue(Text);
            if (SharedString == null)
                SharedString = new Excel.SharedString(Worksheet.Workbook, Text);

            SharedString.AddFontString(font, value);
        }

        /// <summary>
        /// Creates a copy of the Cell and assigns to the given Row and Column.
        /// </summary>
        public Cell Clone(Row row, Column column)
        {
            Cell newCell = new Cell(row, column, Value, StyleIndex);
            return newCell;
        }

        /// <summary>
        /// Creates a copy of the Cell and assigns to the given Row.
        /// </summary>
        public Cell Clone(Row row)
        {
            return Clone(row, row.Worksheet.Columns[Column.Index]);           
        }

        /// <summary>
        /// Creates a copy of the Cell and assigns to the given Column.
        /// </summary>
        public Cell Clone(Column column)
        {
            return Clone(column.Worksheet.Rows[Row.Index], column);
        }

        /// <summary>
        /// Creates a copy of the Cell.
        /// </summary>
        public Cell Clone()
        {
            return Clone(Row, Column);
        }

        /// <summary>
        /// Copies and pastes the Cell to the given address on the given Worksheet. Overwrites existing target Cell.
        /// </summary>
        public Cell CopyTo(Worksheet worksheet, int rowIndex, int columnIndex)
        {
            return CopyTo(worksheet.Rows[rowIndex], worksheet.Columns[columnIndex]);
        }

        /// <summary>
        /// Copies and pastes the Cell to the given address. Overwrites existing target Cell.
        /// </summary>
        public Cell CopyTo(Row row, Column column)
        {
            if (row.Worksheet != column.Worksheet)
                throw new Exception("Error in Cell.CopyTo(row, column). Row and column worksheets do not match");

            Cell newCell = Clone(row, column);
            row.Worksheet.Cells[row.Index, column.Index] = newCell;
            return newCell;
        }

        /// <summary>
        /// Copies and pastes the Cell to the given address on the same Worksheet. Overwrites existing target Cell.
        /// </summary>
        public Cell CopyTo(int rowIndex, int columnIndex)
        {
            return CopyTo(Worksheet, rowIndex, columnIndex);
        }

        /// <summary>
        /// Clears the Cell value and format.
        /// </summary>
        public void Clear()
        {
            UpdateCellValue("");
            base.StyleIndex = CellFormat.DefaultStyleIndex;
        }

        /// <summary>
        /// Deletes the Cell from its Row and shifts all right hand side cells to the left by one.
        /// </summary>
        public void Delete()
        {
            Row.Cells.Delete(Column.Index);
        }

        /// <summary>
        /// Gets the Cell value as a DateTime value. Returns DateTime.MinValue if the Cell value isn't a valid number.
        /// </summary>
        /// <returns></returns>
        public DateTime TryGetDateTimeCellValue()
        {            
            double doubleValue = 0;
            if (double.TryParse(Value.ToString(), out doubleValue))
                return ExcelUtilities.GetTimeConverter().GetDateTimeValueFromDoubleValue(doubleValue);
            else
                return DateTime.MinValue;
        }

        /***********************************
         * STATIC METHODS
         ************************************/


        /// <summary>
        /// Gets a Cell A1 style address for the given (1,1) style address.
        /// </summary>
        public static string GetCellAddressByColumnNameByRowIndex(string columnName, int rowIndex)
        {
            return string.Concat(columnName, rowIndex);
        }

        /***********************************
         * INTERNAL METHODS
         ************************************/


        internal void AddToRowCollection()
        {
            Row.Cells.AddToCollection(this);
        }

        /// <summary>
        /// Just removes the Cell from its Row rather than performing a full Delete.
        /// </summary>
        internal void RemoveFromRowCollection()
        {
            Row.Cells.Remove(Column.Index);
        }


        /***********************************
         * PRIVATE METHODS
         ************************************/

        private static bool CellValueIsNumeric(object value)
        {
            double doubleValue;
            return double.TryParse(value.ToString(), out doubleValue);
        }

        private object GetValue()
        {
            switch (CellDataType)
            {
                case CellDataType.Blank:
                    return "";
                case CellDataType.Boolean:
                    return (bool)_value;
                case CellDataType.SharedString:
                    return _value.ToString();
                case CellDataType.Number:
                    return double.Parse(_value.ToString());
                case CellDataType.String:
                    return (string)_value;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }        

        private void UpdateCellValue(object value)
        {
            SharedString = null;

            CellDataType cellDataType = CellDataType.Blank;
            
            if (value == null)
                value = "";

            if (value.ToString() == "")
            {
                cellDataType = CellDataType.Blank;
            }
            else if (value.GetType() == typeof(bool))
            {
                cellDataType = CellDataType.Boolean;
            }
            else if (value.GetType() == typeof(DateTime))
            {
                value = ExcelUtilities.GetTimeConverter().GetDoubleValueFromDateTimeValue((DateTime)value);
                cellDataType = CellDataType.Number; // Date DataType is new in 2010. Don't support yet.
            }
            else if (value.GetType() == typeof(string))
            {              
                cellDataType = CellDataType.SharedString;
            }
            else if (CellValueIsNumeric(value))
            {
                cellDataType = CellDataType.Number;
            }
            else
            {
                throw new ArgumentOutOfRangeException();
            }

            _value = value;
            CellDataType = cellDataType;

            ResetRequiredDimensions();
        }

        private void ResetRequiredDimensions()
        {
            // should also be called when the font is updated. Not currently handled
            _requiredHeight = -1;
            _requiredWidth = -1;
        }

        private void UpdateHyperlink(Cell targetCell)
        {
            _hyperlinkToCell = targetCell;

            if (targetCell != null)
            {
                Font.Color.Rgb = Color.GetRgb(0, 0, 255);
                Font.Underline = true;
            }
        }





        /***********************************
         * DAL METHODS
         ************************************/

        
        // Read
        internal static Cell ReadCellFromReader(CustomOpenXmlReader reader, Row row)
        {
            CellDataType cellDataType = CellDataType.Number;
            int styleIndex = CellFormat.DefaultStyleIndex;
            string address = "";
            object value = null;

            foreach (CustomOpenXmlAttribute attribute in reader.Attributes)
            {
                switch (attribute.LocalName)
                {
                    case "t":
                        cellDataType = GetCellDataTypeFromAttributeValue(attribute.Value);
                        break;
                    case "s":
                        styleIndex = attribute.GetIntValue();
                        break;
                    case "r":
                        address = attribute.Value;
                        break;

                }
            }

            value = GetCellValueFromReader(reader, cellDataType, row);
            if (value == null)
                cellDataType = CellDataType.Blank;

            // Address doesn't technically need to be included. Needs handling, but keep simple for now.
            int columnIndex = ExcelUtilities.GetColumnIndexConverter().GetColumnIndexFromCellAddress(address);

            Cell cell = new Cell(row, row.Worksheet.Columns[columnIndex], value, cellDataType, styleIndex);
            return cell;
        }

        private static CellDataType GetCellDataTypeFromAttributeValue(string attributeValue)
        {
            switch (attributeValue)
            {
                case "b":
                    return CellDataType.Boolean;
                case "s":
                    return CellDataType.SharedString;
                case "str":
                    return CellDataType.String;
                default:
                    return CellDataType.Number;
            }
        }
        private static object GetCellValueFromReader(CustomOpenXmlReader reader, CellDataType cellDataType, Row row)
        {
            object cellValue = null;

            while (reader.ReadToEndElement<OpenXmlSpreadsheet.Cell>())
            {
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.CellValue>())
                {
                    string rawValue = reader.GetText();

                    if (rawValue != "")
                    {
                        switch (cellDataType)
                        {
                            case CellDataType.Boolean:
                                cellValue = (rawValue == "1" ? true : false);
                                break;
                            case CellDataType.SharedString:
                                int sharedStringIndex = int.Parse(rawValue);
                                cellValue = row.Worksheet.Workbook.SharedStrings[sharedStringIndex];
                                break;
                            case CellDataType.Number:
                                double doubleValue = 0;
                                if (double.TryParse(rawValue, out doubleValue))
                                    cellValue = doubleValue;
                                break;
                            case CellDataType.String:
                                cellValue = rawValue.ToString();
                                break;
                            default:
                                throw new ArgumentOutOfRangeException();
                        }
                    }
                }
            }

            return cellValue;
        }

        // Write
        internal static void WriteCellToWriter(CustomOpenXmlWriter<OpenXmlPackaging.WorksheetPart> writer, Cell cell)
        {
            if (!cell.IsUsed)
                return;
            
            writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.Cell());

            string dataType = GetAttributeValueFromCellDataType(cell.CellDataType);
            if (dataType != "") writer.WriteAttribute("t", dataType);
            if (cell.StyleIndex > CellFormat.DefaultStyleIndex) writer.WriteAttribute("s", cell.StyleIndex);
            writer.WriteAttribute("r", cell.Address);

            if (cell.RawValue != null && cell.RawValue.ToString() != "")
            {
                writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.CellValue());

                switch (cell.CellDataType)
                {
                    case CellDataType.Boolean:
                        writer.WriteText((bool)cell.RawValue);
                        break;
                    case CellDataType.SharedString:
                        int sharedStringIndex;
                        if (cell.SharedString != null)
                            sharedStringIndex = cell.Worksheet.Workbook.SharedStrings[cell.SharedString];
                        else
                            sharedStringIndex = cell.Worksheet.Workbook.SharedStrings[cell.RawValue.ToString()];

                        writer.WriteText(sharedStringIndex.ToString());
                        break;
                    default:
                        writer.WriteString(cell.RawValue.ToString());
                        break;
                }
                writer.WriteEndElement();    // CellValue
            }

            writer.WriteEndElement();   // Cell
        }

        private static string GetAttributeValueFromCellDataType(CellDataType cellDataType)
        {
            switch (cellDataType)
            {
                case CellDataType.Number:
                case CellDataType.Blank:
                    return "";
                default:
                    return new OpenXml.EnumValue<OpenXmlSpreadsheet.CellValues>((OpenXmlSpreadsheet.CellValues)((int)cellDataType));
            }
        }

    }
}
