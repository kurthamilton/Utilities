using System;
using System.Collections.Generic;

using System.Xml;
using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{
    // keep internal for now until fully developed
    public class CellFormat : BaseExcel, IEquatable<CellFormat>
    {
        internal const int DefaultBaseFormatId = 0;
        internal const int DefaultStyleIndex = 0;

        private Styles _styles;
        internal Styles Styles { get { if (_styles == null) return BaseRange.Worksheet.Workbook.Styles; return _styles; } private set { _styles = value; } }
        private BaseRange BaseRange { get; set; }

        private int _cellFormatId = DefaultStyleIndex;
        internal int CellFormatId { get { return _cellFormatId; } private set { _cellFormatId = value; } }

        internal int BaseFormatId { get; private set; }        

        private Alignment _alignment;
        public Alignment Alignment { get { return GetAlignment();  } private set { _alignment = value; } }
        private Borders _borders;
        public Borders Borders { get { return GetBorders(); } private set { _borders = value; } }
        private Fill _fill;
        public Fill Fill { get { return GetFill(); } private set { _fill = value; } }
        private Font _font;
        public Font Font { get { return GetFont(); } private set { _font = value; } }
        private NumberFormat _numberFormat;
        public NumberFormat NumberFormat { get { return GetNumberFormat(); } private set { _numberFormat = value; } }        

        /***********************************
         * CONTRUCTORS
         ************************************/

        // range level cell format - used to update range cell format properties
        internal CellFormat(BaseRange range)
        {
            BaseRange = range;
        }

        // workbook level cell formats - used to define workbook cell formats
        internal CellFormat(CellFormat cellFormat)
            : this(cellFormat.Styles, cellFormat.Alignment, cellFormat.Borders, cellFormat.Fill, cellFormat.Font, cellFormat.NumberFormat)
        {
        }

        internal CellFormat(Styles styles)
            : this (styles, DefaultBaseFormatId, DefaultStyleIndex, null, null, null, null, new NumberFormat(styles))
        {
        }

        internal CellFormat(Styles styles, int baseFormatId, int cellFormatId, Alignment alignment, Borders borders, 
            Fill fill, Font font, NumberFormat numberFormat)
        {
            Styles = styles;

            BaseFormatId = baseFormatId;
            CellFormatId = cellFormatId;
            Alignment = alignment;
            Borders = borders;
            Fill = fill;
            Font = font;
            NumberFormat = numberFormat;            
        }

        internal CellFormat(Styles styles, Alignment alignment, Borders borders, Fill fill, Font font, NumberFormat numberFormat)
            : this(styles, DefaultBaseFormatId, DefaultStyleIndex, alignment, borders, fill, font, numberFormat)
        {
        }

        /***********************************
         * PUBLIC METHODS
         ************************************/

        // implement IEquatable
        public bool Equals(CellFormat other)
        {
            return (
                other.Alignment.Equals(Alignment) &&
                other.Borders.Equals(Borders) &&
                other.Fill.Equals(Fill) &&
                other.Font.Equals(Font) &&
                other.NumberFormat.Equals(NumberFormat));
        }

        /***********************************
         * PRIVATE METHODS
         ************************************/
        private Alignment GetAlignment()
        {
            if (_alignment == null)
            {
                if (BaseRange != null)
                    _alignment = new Alignment(BaseRange);
                else
                    _alignment = new Alignment();
            }
            return _alignment;
        }
        private Borders GetBorders()
        {
            if (_borders == null)
            {
                if (BaseRange != null)
                    _borders = new Borders(BaseRange);
                else
                    _borders = Styles.Borders[Borders.DefaultBordersId];
            }
            return _borders;
        }
        private Fill GetFill()
        {
            if (_fill == null)
            {
                if (BaseRange != null)
                    _fill = new Fill(BaseRange);
                else
                    _fill = Styles.Fills[Fill.DefaultFillId];
            }
            return _fill;
        }
        private Font GetFont()
        {
            if (_font == null)
            {
                if (BaseRange != null)
                    _font = new Font(BaseRange);
                else
                    _font = Styles.Fonts[Font.DefaultFontId];
            }
            return _font;
        }
        private NumberFormat GetNumberFormat()
        {
            if (_numberFormat == null)
            {
                if (BaseRange != null)
                    _numberFormat = new NumberFormat(BaseRange);
                else
                    _numberFormat = Styles.NumberFormats[NumberFormat.DefaultNumberFormatId];
            }
            return _numberFormat;
        }

        /***********************************
         * DAL METHODS
         ************************************/

        // Read

        internal static CellFormat ReadCellFormatFromReader(CustomOpenXmlReader reader, Styles styles, int cellFormatId)
        {
            CellFormat cellFormat = new CellFormat(styles);
            cellFormat.CellFormatId = cellFormatId;

            //CustomOpenXmlAttribute baseFormatAttribute = reader.Attributes.GetAttribute("xfId");
            //if (baseFormatAttribute != null)
            //    cellFormat.BaseFormatId = baseFormatAttribute.GetIntValue();

            bool applyAlignment = false;
            if (reader.Attributes.TryGetBoolAttributeValue("applyNumberFormat"))
                cellFormat.NumberFormat = styles.NumberFormats[reader.Attributes["numFmtId"].GetIntValue()];
            if (reader.Attributes.TryGetBoolAttributeValue("applyBorder"))
                cellFormat.Borders = styles.Borders[reader.Attributes["borderId"].GetIntValue()];
            if (reader.Attributes.TryGetBoolAttributeValue("applyFill"))
                cellFormat.Fill = styles.Fills[reader.Attributes["fillId"].GetIntValue()];
            if (reader.Attributes.TryGetBoolAttributeValue("applyFont"))
                cellFormat.Font = styles.Fonts[reader.Attributes["fontId"].GetIntValue()];
            if (reader.Attributes.TryGetBoolAttributeValue("applyAlignment"))
                applyAlignment = true;

            while (reader.ReadToEndElement<OpenXmlSpreadsheet.CellFormat>())
            {
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.Alignment>() && applyAlignment)
                {
                    cellFormat.Alignment = Alignment.ReadAlignmentFromReader(reader);
                }
            }

            return cellFormat;
        }

        // Write

        internal static void WriteCellFormatToWriter(CustomOpenXmlWriter<OpenXmlPackaging.WorkbookStylesPart> writer, CellFormat cellFormat)
        {
            writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.CellFormat());

            if (cellFormat.NumberFormat != null && cellFormat.NumberFormat.NumberFormatId != NumberFormat.DefaultNumberFormatId)
            {
                writer.WriteAttribute("applyNumberFormat", true);
                writer.WriteAttribute("numFmtId", cellFormat.NumberFormat.NumberFormatId);
            }
            if (cellFormat._borders != null)
            {
                writer.WriteAttribute("applyBorder", true);
                writer.WriteAttribute("borderId", cellFormat.Borders.BordersId);
            }
            if (cellFormat._fill != null)
            {
                writer.WriteAttribute("applyFill", true);
                writer.WriteAttribute("fillId", cellFormat.Fill.FillId);
            }
            if (cellFormat._font != null)
            {
                writer.WriteAttribute("applyFont", true);
                writer.WriteAttribute("fontId", cellFormat.Font.FontId);
            }
            if (cellFormat._alignment != null)
            {
                writer.WriteAttribute("applyAlignment", true);
                Alignment.WriteAlignmentToWriter(writer, cellFormat.Alignment);
            }

            writer.WriteEndElement();   // CellFormat
        }
    }
}
