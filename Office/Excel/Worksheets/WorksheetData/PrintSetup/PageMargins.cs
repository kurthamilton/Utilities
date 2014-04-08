using System;
using System.Collections.Generic;

using System.Xml;
using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{
    public class PageMargins : BaseExcel
    {
        public Worksheet Worksheet { get; private set; }
        
        // Need to work out relationship between value stored here and value presented in UI.
        private double _left = 0.6;
        public double Left { get { return _left; } set { if (value >= 0) _left = value; } }
        private double _right = 0.6;
        public double Right { get { return _right; } set { if (value >= 0) _right = value; } }
        private double _top = 0.8;
        public double Top { get { return _top; } set { if (value >= 0) _top = value; } }
        private double _bottom = 0.8;
        public double Bottom { get { return _bottom; } set { if (value >= 0) _bottom = value; } }
        private double _header = 0.4;
        public double Header { get { return _header; } set { if (value >= 0) _header = value; } }
        private double _footer = 0.4;
        public double Footer { get { return _footer; } set { if (value >= 0) _footer = value; } }


        /***********************************
         * CONSTRUCTORS
         ************************************/


        public PageMargins(Worksheet worksheet)
        {
            Worksheet = worksheet;
        }
        public PageMargins(Worksheet worksheet, double left, double right, double top, double bottom, double header, double footer)
            : this (worksheet)
        {
            Left = left;
            Right = right;
            Top = top;
            Bottom = bottom;
            Header = header;
            Footer = footer;
        }

        /***********************************
         * INTERNAL METHODS
         ************************************/


        internal PageMargins Clone(Worksheet worksheet)
        {
            PageMargins newPageMargins = new PageMargins(worksheet,
                Left, Right, Top, Bottom, Header, Footer);
            return newPageMargins;
        }

        /***********************************
         * DAL METHODS
         ************************************/
        
        // Read

        internal static PageMargins ReadPageMarginsFromReader(CustomOpenXmlReader reader, Worksheet worksheet)
        {
            PageMargins pageMargins = new PageMargins(worksheet);

            foreach (CustomOpenXmlAttribute attribute in reader.Attributes)
            {
                switch (attribute.LocalName)
                {
                    case "left":
                        pageMargins.Left = attribute.GetDoubleValue();
                        break;
                    case "right":
                        pageMargins.Right = attribute.GetDoubleValue();
                        break;
                    case "top":
                        pageMargins.Top = attribute.GetDoubleValue();
                        break;
                    case "bottom":
                        pageMargins.Bottom = attribute.GetDoubleValue();
                        break;
                    case "header":
                        pageMargins.Header = attribute.GetDoubleValue();
                        break;
                    case "footer":
                        pageMargins.Footer = attribute.GetDoubleValue();
                        break;
                    default:
                        throw new Exception(string.Format("PageMargins attribute {0} not coded", attribute.LocalName));
                }
            }

            return pageMargins;
        }

        // Write

        internal static void WritePageMarginsToWriter(CustomOpenXmlWriter<OpenXmlPackaging.WorksheetPart> writer, PageMargins pageMargins)
        {
            writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.PageMargins());

            writer.WriteAttribute("left", pageMargins.Left);
            writer.WriteAttribute("right", pageMargins.Right);
            writer.WriteAttribute("top", pageMargins.Top);
            writer.WriteAttribute("bottom", pageMargins.Bottom);
            writer.WriteAttribute("header", pageMargins.Header);
            writer.WriteAttribute("footer", pageMargins.Footer);

            writer.WriteEndElement();   // PageMargins
        }

    }
}
