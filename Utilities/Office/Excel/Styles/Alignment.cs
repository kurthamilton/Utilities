using System;
using System.Collections.Generic;

using System.Xml;
using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{
    public class Alignment : BaseExcel, IEquatable<Alignment>
    {
        private BaseRange BaseRange { get; set; }

        private bool _wrapText;
        public bool WrapText { get { return _wrapText; } set { _wrapText = value; UpdateProperty(); } }
        private HorizontalAlignmentOptions _horizontal;
        public HorizontalAlignmentOptions Horizontal { get { return _horizontal; } set { _horizontal = value; UpdateProperty(); } }
        private VerticalAlignmentOptions _vertical;
        public VerticalAlignmentOptions Vertical { get { return _vertical; } set { _vertical = value; UpdateProperty(); } }        

        /***********************************
         * CONTRUCTORS
         ************************************/

        internal Alignment(BaseRange range)
            : this(range.Worksheet.Workbook.Styles.CellFormats[range.StyleIndex].Alignment)
        {
            BaseRange = range;
        }

        internal Alignment(Alignment alignment)
            : this(alignment.WrapText, alignment.Horizontal, alignment.Vertical)
        {
        }        

        internal Alignment(bool wrapText = false, 
            HorizontalAlignmentOptions horizontal = (HorizontalAlignmentOptions)0,
            VerticalAlignmentOptions vertical = (VerticalAlignmentOptions)0)
        {
            WrapText = wrapText;
            Vertical = vertical;
            Horizontal = horizontal;
        }

        private Alignment()
        {
        }

        /***********************************
         * PUBLIC METHODS
         ************************************/

        // implement IEquatable
        public bool Equals(Alignment other)
        {
            return (other.Horizontal == Horizontal && other.Vertical == Vertical && other.WrapText == WrapText);
        }


        /***********************************
         * PRIVATE METHODS
         ************************************/

        private void UpdateProperty()
        {
            if (BaseRange != null)
            {
                // update base range with new/existing cell format id             
                Styles styles = BaseRange.Worksheet.Workbook.Styles;
                CellFormat cellFormat = styles.CellFormats[BaseRange.StyleIndex];
                CellFormat newCellFormat = styles.CellFormats.Insert(
                    new CellFormat(cellFormat.Styles, this, cellFormat.Borders, cellFormat.Fill, cellFormat.Font, cellFormat.NumberFormat));
                BaseRange.StyleIndex = newCellFormat.CellFormatId;
            }
        }

        /***********************************
         * DAL METHODS
         ************************************/

        // Read

        internal static Alignment ReadAlignmentFromReader(CustomOpenXmlReader reader)
        {
            Alignment alignment = new Alignment();

            foreach (CustomOpenXmlAttribute attribute in reader.Attributes)
            {
                switch (attribute.LocalName)
                {
                    case "wrapText":
                        alignment.WrapText = attribute.GetBoolValue();
                        break;
                    case "horizontal":
                        alignment.Horizontal = Helpers.GetEnumValueFromDescription<HorizontalAlignmentOptions>(attribute.Value);
                        break;
                    case "vertical":
                        alignment.Vertical = Helpers.GetEnumValueFromDescription<VerticalAlignmentOptions>(attribute.Value);
                        break;
                }
            }

            return alignment;
        }

        // Write

        internal static void WriteAlignmentToWriter(CustomOpenXmlWriter<OpenXmlPackaging.WorkbookStylesPart> writer, Alignment alignment)
        {
            writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.Alignment());

            if (alignment.WrapText) writer.WriteAttribute("wrapText", true);
            if ((int)alignment.Horizontal > 0) writer.WriteAttribute("horizontal", Helpers.ToCamelCase(alignment.Horizontal.ToString()));
            if ((int)alignment.Vertical > 0) writer.WriteAttribute("vertical", Helpers.ToCamelCase(alignment.Vertical.ToString()));

            writer.WriteEndElement();   // Alignment
        }
    }
}
