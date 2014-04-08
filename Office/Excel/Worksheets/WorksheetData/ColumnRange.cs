using System;
using System.Collections.Generic;

using System.Xml;
using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{
    internal class ColumnRange : BaseExcel
    {        
        public Worksheet Worksheet { get; private set; }
        public int MinIndex { get; internal set; }
        public int MaxIndex { get; internal set; }
        
        public int StyleIndex { get; private set; }
        public double Width { get; private set; }
        
        public bool IsBestFit { get; private set; }
        public bool IsCustomWidth { get; private set; }

        internal const double DefaultColumnWidth = 9.140625;

        /***********************************
         * CONSTRUCTORS
         ************************************/

        internal ColumnRange(Column startColumn, Column endColumn)
        {
            Worksheet = startColumn.Worksheet;
            MinIndex = startColumn.Index;
            MaxIndex = endColumn.Index;
            StyleIndex = startColumn.StyleIndex;
            Width = startColumn.Width;
            IsBestFit = startColumn.IsBestFit;
            IsCustomWidth = startColumn.IsCustomWidth;
        }

        // internal for time being - until full styling is required
        internal ColumnRange(Worksheet worksheet, int minIndex, int maxIndex, int styleIndex, double width, bool isBestFit, bool isCustomWidth)
        {
            Worksheet = worksheet;
            MinIndex = minIndex;
            MaxIndex = maxIndex;            
            StyleIndex = styleIndex;
            Width = width;
            IsBestFit = isBestFit;
            IsCustomWidth = isCustomWidth;
        }


        /***********************************
         * PUBLIC METHODS
         ************************************/


        /// <summary>
        /// Creates a copy of the ColumnRange and assigns to the given Worksheet.
        /// </summary>
        public ColumnRange Clone(Worksheet worksheet)
        {
            ColumnRange newColumnRange = new ColumnRange(worksheet, MinIndex, MaxIndex, StyleIndex, Width, IsBestFit, IsCustomWidth);
            return newColumnRange;
        }



        /***********************************
         * INTERNAL METHODS
         ************************************/

        internal static List<ColumnRange> GetColumnRangesFromWorksheet(Worksheet worksheet)
        {
            List<ColumnRange> columnRanges = new List<ColumnRange>();
            Column startColumn = null;
            Column prevColumn = null;
            foreach (Column column in worksheet.Columns)
            {
                if (prevColumn == null)
                {
                    startColumn = column;
                    prevColumn = column;
                }
                else
                {
                    if (column.Similar(prevColumn))
                    {
                        prevColumn = column;
                    }
                    else
                    {
                        columnRanges.Add(new ColumnRange(startColumn, prevColumn));
                        prevColumn = column;
                        startColumn = column;
                    }
                }
            }
            columnRanges.Add(new ColumnRange(startColumn, prevColumn));

            return columnRanges;
        }


        /***********************************
         * PRIVATE METHODS
         ************************************/



        /***********************************
         * DAL METHODS
         ************************************/


        // Read
        
        internal static ColumnRange ReadColumnRangeFromReader(CustomOpenXmlReader reader, Worksheet worksheet)
        {
            int min = 0;
            int max = 0;
            int styleIndex = CellFormat.DefaultStyleIndex;
            double width = -1;
            bool isBestFit = false;
            bool isCustomWidth = false;

            foreach (CustomOpenXmlAttribute attribute in reader.Attributes)
            {
                switch (attribute.LocalName)
                {
                    case "min":
                        min = attribute.GetIntValue();
                        break;
                    case "max":
                        max = attribute.GetIntValue();
                        break;
                    case "style":
                        styleIndex = attribute.GetIntValue();
                        break;
                    case "width":
                        width = attribute.GetDoubleValue();
                        break;
                    case "bestFit":
                        isBestFit = attribute.GetBoolValue();
                        break;
                    case "customWidth":
                        isCustomWidth = attribute.GetBoolValue();
                        break;
                }
            }

            ColumnRange columnRange = new ColumnRange(worksheet, min, max, styleIndex, width, isBestFit, isCustomWidth);
            return columnRange;
        }

        // Write
        
        internal static void WriteColumnRangeToWriter(CustomOpenXmlWriter<OpenXmlPackaging.WorksheetPart> writer, ColumnRange columnRange)
        {
            writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.Column());

            writer.WriteAttribute("min", columnRange.MinIndex);
            writer.WriteAttribute("max", columnRange.MaxIndex);
            if (columnRange.StyleIndex > 0) writer.WriteAttribute("style", columnRange.StyleIndex);
            writer.WriteAttribute("width", columnRange.Width);
            if (columnRange.IsBestFit) writer.WriteAttribute("bestFit", columnRange.IsBestFit);
            if (columnRange.IsCustomWidth) writer.WriteAttribute("customWidth", columnRange.IsCustomWidth);
            
            writer.WriteEndElement();   // Column
        }
    }
}
