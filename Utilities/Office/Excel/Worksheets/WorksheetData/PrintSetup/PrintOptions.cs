using System;
using System.Collections.Generic;

using System.Xml;
using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{
    public class PrintOptions : BaseExcel
    {
        public Worksheet Worksheet { get; private set; }
        
        // printOptions
        public bool HorizontalCentered { get; set; }
        public bool VerticalCentered { get; set; }
        public bool GridLines { get; set; }
        public bool Headings { get; set; }

        /***********************************
         * CONSTRUCTORS
         ************************************/

        public PrintOptions(Worksheet worksheet)
        {
            Worksheet = worksheet;
        }

        public PrintOptions(Worksheet worksheet, bool horizontalCentered = false, bool verticalCentered = false, bool gridLines = false, bool headings = false)
            : this (worksheet)
        {
            HorizontalCentered = horizontalCentered;
            VerticalCentered = verticalCentered;
            GridLines = gridLines;
            Headings = headings;
        }

        /***********************************
         * INTERNAL METHODS
         ************************************/

        internal PrintOptions Clone(Worksheet worksheet)
        {
            PrintOptions newPrintOptions = new PrintOptions(worksheet, HorizontalCentered, VerticalCentered, GridLines, Headings);
            return newPrintOptions;
        }

        internal bool HasValue()
        {
            return (HorizontalCentered || VerticalCentered || GridLines || Headings);
        }

        /***********************************
         * DAL METHODS
         ************************************/

        // Read

        internal static PrintOptions ReadPrintOptionsFromReader(CustomOpenXmlReader reader, Worksheet worksheet)
        {
            PrintOptions printOptions = new PrintOptions(worksheet);

            foreach (CustomOpenXmlAttribute attribute in reader.Attributes)
            {
                switch (attribute.LocalName)
                {
                    case "horizontalCentered":
                        printOptions.HorizontalCentered = attribute.GetBoolValue();
                        break;
                    case "verticalCentered":
                        printOptions.VerticalCentered = attribute.GetBoolValue();
                        break;
                    case "gridLines":
                        printOptions.GridLines = attribute.GetBoolValue();
                        break;
                    case "headings":
                        printOptions.Headings = attribute.GetBoolValue();
                        break;
                    default:
                        throw new Exception(string.Format("PrintOptions attribute {0} not coded", attribute.LocalName));
                }
            }

            return printOptions;
        }


        // Write

        internal static void WritePrintOptionsToWriter(CustomOpenXmlWriter<OpenXmlPackaging.WorksheetPart> writer, PrintOptions printOptions)
        {
            if (printOptions.HasValue())
            {
                writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.PrintOptions());

                if (printOptions.HorizontalCentered) writer.WriteAttribute("horizontalCentered", printOptions.HorizontalCentered);
                if (printOptions.VerticalCentered) writer.WriteAttribute("verticalCentered", printOptions.VerticalCentered);
                if (printOptions.GridLines) writer.WriteAttribute("gridLines", printOptions.GridLines);
                if (printOptions.Headings) writer.WriteAttribute("headings", printOptions.Headings);

                writer.WriteEndElement();   // PrintOptions
            }
        }
    }
}
