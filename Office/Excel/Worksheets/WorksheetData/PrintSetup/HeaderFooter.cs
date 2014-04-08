using System;
using System.Collections.Generic;

using System.Xml;
using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{
    public class HeaderFooter : BaseExcel
    {
        public Worksheet Worksheet { get; private set; }

        public string OddHeader { get; set; }
        public string OddFooter { get; set; }

        private bool _alignWithMargins = true;
        public bool AlignWithMargins { get { return _alignWithMargins; } set { _alignWithMargins = value; } }
        // TO BE DEVELOPED - evenHeader, evenFooter, firstHeader, firstFooter, differentFirst, differentOddEven, scaleWithDoc


        /***********************************
         * CONSTRUCTORS
         ************************************/


        public HeaderFooter(Worksheet worksheet)
        {
            Worksheet = worksheet;
        }

        public HeaderFooter(Worksheet worksheet, string oddHeader, string oddFooter, bool alignHeaderFooterWithMargins)
            : this (worksheet)
        {
            OddHeader = oddHeader;
            OddFooter = oddFooter;
            AlignWithMargins = alignHeaderFooterWithMargins;
        }


        /***********************************
         * INTERNAL METHODS
         ************************************/

        internal HeaderFooter Clone(Worksheet worksheet)
        {
            HeaderFooter newHeaderFooter = new HeaderFooter(worksheet, OddHeader, OddFooter, AlignWithMargins);
            return newHeaderFooter;
        }

        internal bool HasValue()
        {
            return (OddHeader != "" ||
                OddFooter != "" ||
                !AlignWithMargins);
        }

        /***********************************
         * DAL METHODS
         ************************************/

        // Read

        internal static HeaderFooter ReadHeaderFooterFromReader(CustomOpenXmlReader reader, Worksheet worksheet)
        {
            HeaderFooter headerFooter = new HeaderFooter(worksheet);

            foreach (CustomOpenXmlAttribute attribute in reader.Attributes)
            {
                switch (attribute.LocalName)
                {
                    case "alignWithMargins":
                        headerFooter.AlignWithMargins = attribute.GetBoolValue();
                        break;
                    default:
                        throw new Exception(string.Format("HeaderFooter attribute {0} not coded", attribute.LocalName));
                }
            }

            while (reader.ReadToEndElement<OpenXmlSpreadsheet.HeaderFooter>())
            {
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.OddHeader>())
                    headerFooter.OddHeader = reader.GetText();
                else if (reader.IsStartElementOfType<OpenXmlSpreadsheet.OddFooter>())
                    headerFooter.OddFooter = reader.GetText();
            }

            return headerFooter;
        }

        // Write

        internal static void WriteHeaderFooterToWriter(CustomOpenXmlWriter<OpenXmlPackaging.WorksheetPart> writer, HeaderFooter headerFooter)
        {
            if (headerFooter.HasValue())
            {
                writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.HeaderFooter());
                if (!headerFooter.AlignWithMargins) writer.WriteAttribute("alignWithMargins", headerFooter.AlignWithMargins);

                if (headerFooter.OddHeader != "")
                {
                    writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.OddHeader());
                    writer.WriteText(headerFooter.OddHeader);
                    writer.WriteEndElement();   // OddHeader
                }
                if (headerFooter.OddFooter != "")
                {
                    writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.OddFooter());
                    writer.WriteText(headerFooter.OddFooter);
                    writer.WriteEndElement();   // OddFooter
                }

                writer.WriteEndElement();   // HeaderFooter
            }
        }

    }
}
