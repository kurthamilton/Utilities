using System;
using System.Collections.Generic;
using System.Linq;

using System.Xml;
using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{    
    public class Borders : BaseExcel, IEquatable<Borders>
    {
        internal const int DefaultBordersId = 0;

        internal Styles Styles { get; private set; }
        private BaseRange BaseRange { get; set; }

        private int _bordersId = DefaultBordersId;
        internal int BordersId { get { return _bordersId; } private set { _bordersId = value; } }

        private Border _left;
        public Border Left { get { return GetBorder(BorderType.Left); } internal set { _left = value; } }
        private Border _right;
        public Border Right { get { return GetBorder(BorderType.Right); } internal set { _right = value; } }
        private Border _top;
        public Border Top { get { return GetBorder(BorderType.Top); } internal set { _top = value; } }
        private Border _bottom;
        public Border Bottom { get { return GetBorder(BorderType.Bottom); } internal set { _bottom = value; } }
        private Border _diagonal;
        public Border Diagonal { get { return GetBorder(BorderType.Diagonal); } internal set { _diagonal = value; } }

        // range level borders - used to update range border properties
        internal Borders(BaseRange range)
            : this(range.Worksheet.Workbook.Styles.CellFormats[range.StyleIndex].Borders)
        {
            BaseRange = range;
        }
        
        internal Borders(Borders borders)
            : this(borders.Styles, DefaultBordersId)
        {            
        }

        // workbook level borders - used to define workbook borders
        internal Borders(Styles styles, int bordersId)
            : this (styles, bordersId, null, null, null, null, null)
        {
            
        }

        internal Borders(Styles styles, int bordersId, Border left, Border right, Border top, Border bottom, Border diagonal)
        {
            Styles = styles;
            BordersId = bordersId;
            Left = left;
            Right = right;
            Top = top;
            Bottom = bottom;
            Diagonal = diagonal;
        }

        /***********************************
         * INTERNAL PROPERTIES
         ************************************/


        /***********************************
         * PUBLIC METHODS
         ************************************/

        // implement IEquatable
        public bool Equals(Borders other)
        {
            return (other.Left.Equals(Left) && 
                other.Right.Equals(Right) && 
                other.Top.Equals(Top) && 
                other.Bottom.Equals(Bottom) && 
                other.Diagonal.Equals(Diagonal));
        }

        /***********************************
         * INTERNAL METHODS
         ************************************/
        internal void SetBorder(Border value)
        {
            if (value != null)
            {
                switch (value.Type)
                {
                    case BorderType.Left:
                        _left = value;
                        break;
                    case BorderType.Right:
                        _right = value;
                        break;
                    case BorderType.Top:
                        _top = value;
                        break;
                    case BorderType.Bottom:
                        _bottom = value;
                        break;
                    case BorderType.Diagonal:
                        _diagonal = value;
                        break;
                    default:
                        throw new Exception(string.Format("BorderType {0} not implemented in Borders.SetBorder", value.Type));
                }
            }
        }

        internal void UpdateBorder(Border value)
        {
            SetBorder(value);

            if (BaseRange != null)
            {
                // update base range with new/existing cell format id
                Borders newBorders = Styles.Borders.Insert(this);

                CellFormat cellFormat = Styles.CellFormats[BaseRange.StyleIndex];
                CellFormat newCellFormat = Styles.CellFormats.Insert(
                    new CellFormat(cellFormat.Styles, cellFormat.Alignment, newBorders, cellFormat.Fill, cellFormat.Font, cellFormat.NumberFormat));
                BaseRange.StyleIndex = newCellFormat.CellFormatId;
            }
        }
       
        /***********************************
         * PRIVATE METHODS
         ************************************/        

        private Border GetBorder(BorderType type)
        {
            Border border;
            switch (type)
            {
                case BorderType.Left:
                    border = _left;
                    break;
                case BorderType.Right:
                    border = _right;
                    break;
                case BorderType.Top:
                    border = _top;
                    break;
                case BorderType.Bottom:
                    border = _bottom;
                    break;
                case BorderType.Diagonal:
                    border = _diagonal;
                    break;
                default:
                    throw new Exception(string.Format("BorderType {0} not implemented in Borders.GetBorder", type));
            }

            if (border == null)
            {
                if (BaseRange != null)
                    border = new Border(this, BaseRange.Worksheet.Workbook.Styles.CellFormats[BaseRange.StyleIndex].Borders.GetBorder(type));
                else
                    throw new Exception("border not set");
            }

            return border;
        }

        /***********************************
         * DAL METHODS
         ************************************/       

        // Read

        internal static Borders ReadBordersFromReader(CustomOpenXmlReader reader, Styles styles, int bordersId)
        {
            Borders borders = new Borders(styles, bordersId);

            while (reader.ReadToEndElement<OpenXmlSpreadsheet.Border>())
            {
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.LeftBorder>())
                    borders.Left = Border.ReadBorderFromReader(reader, borders);
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.RightBorder>())
                    borders.Right = Border.ReadBorderFromReader(reader, borders);
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.TopBorder>())
                    borders.Top = Border.ReadBorderFromReader(reader, borders);
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.BottomBorder>())
                    borders.Bottom = Border.ReadBorderFromReader(reader, borders);
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.DiagonalBorder>())
                    borders.Diagonal = Border.ReadBorderFromReader(reader, borders);
            }

            return borders;
        }

        // Write

        internal static void WriteBordersToWriter(CustomOpenXmlWriter<OpenXmlPackaging.WorkbookStylesPart> writer, Borders borders)
        {
            writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.Border());

            if (borders._left != null) Border.WriteBorderToWriter(writer, borders.Left, new OpenXmlSpreadsheet.LeftBorder());
            if (borders._right != null) Border.WriteBorderToWriter(writer, borders.Right, new OpenXmlSpreadsheet.RightBorder());
            if (borders._top != null) Border.WriteBorderToWriter(writer, borders.Top, new OpenXmlSpreadsheet.TopBorder());
            if (borders._bottom != null) Border.WriteBorderToWriter(writer, borders.Bottom, new OpenXmlSpreadsheet.BottomBorder());
            if (borders._diagonal != null) Border.WriteBorderToWriter(writer, borders.Diagonal, new OpenXmlSpreadsheet.DiagonalBorder());

            writer.WriteEndElement();   // Border
        }
    }
}
