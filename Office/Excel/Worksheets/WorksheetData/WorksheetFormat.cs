using System;
using System.Collections.Generic;

using System.Xml;
using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{
    public class WorksheetFormat : BaseExcel
    {
        public Worksheet Worksheet { get; private set; }
        public double DefaultRowHeight { get; set; }

        /***********************************
         * CONSTRUCTORS
         ************************************/        


        public WorksheetFormat(Worksheet worksheet)
        {
            Worksheet = worksheet;
        }

        public WorksheetFormat(Worksheet worksheet, double defaultRowHeight)
        {
            Worksheet = worksheet;
            DefaultRowHeight = defaultRowHeight;
        }

        /***********************************
         * INTERNAL METHODS
         ************************************/

        internal WorksheetFormat Clone(Worksheet worksheet)
        {
            WorksheetFormat newFormat = new WorksheetFormat(worksheet, DefaultRowHeight);
            return newFormat;
        }

        internal bool HasValue()
        {
            return DefaultRowHeight > 0;
        }

        /***********************************
         * DAL METHODS
         ************************************/

        // Read

        internal static WorksheetFormat ReadWorksheetFormatFromReader(CustomOpenXmlReader reader, Worksheet worksheet)
        {
            WorksheetFormat worksheetFormat = new WorksheetFormat(worksheet);

            foreach (CustomOpenXmlAttribute attribute in reader.Attributes)
            {
                switch (attribute.LocalName)
                {
                    case "defaultRowHeight":
                        worksheetFormat.DefaultRowHeight = attribute.GetDoubleValue();
                        break;
                }
            }

            return worksheetFormat;
        }

        // Write

        internal static void WriteWorksheetFormatToWriter(CustomOpenXmlWriter<OpenXmlPackaging.WorksheetPart> writer, WorksheetFormat worksheetFormat)
        {
            if (worksheetFormat.HasValue())
            {
                writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.SheetFormatProperties());

                if (worksheetFormat.DefaultRowHeight > 0)
                    writer.WriteAttribute("defaultRowHeight", worksheetFormat.DefaultRowHeight);

                writer.WriteEndElement();   // SheetFormatProperties
            }
        }
    }
}
