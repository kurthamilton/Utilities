using System;
using System.Collections.Generic;

using System.Xml;
using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{
    public class PageSetup : BaseExcel
    {
        public Worksheet Worksheet { get; private set; }
        

        // pageSetup
        public bool BlackAndWhite { get; set; }
        private int _copies = 1;
        public int Copies { get { return _copies; } set { if (value > 0) _copies = value; } }
        public bool Draft { get; set; }
        private int _firstPageNumber = 1;
        public int FirstPageNumber { get { return _firstPageNumber; } set { if (value > 0) _firstPageNumber = value; } }
        private int _fitToHeight = -1;
        public int FitToHeight { get { return _fitToHeight; } set { if (value >= 0) _fitToHeight = value; else _fitToHeight = -1; } }
        private int _fitToWidth = -1;
        public int FitToWidth { get { return _fitToWidth; } set { if (value >= 0) _fitToWidth = value; else _fitToHeight = -1; } }
        private int _horizontalDpi = 600;
        public int HorizontalDpi { get { return _horizontalDpi; } set { if (value > 0) _horizontalDpi = value; } }
        private int _verticalDpi = 600;
        public int VerticalDpi { get { return _verticalDpi; } set { if (value > 0) _verticalDpi = value; } }
        internal string RelationshipId { get; set; }
        private WorksheetOrientation _orientation = WorksheetOrientation.Portrait;
        public WorksheetOrientation Orientation { get { return _orientation; } set { if (value != WorksheetOrientation.None) _orientation = value; } }
        private int _paperSize = 9;    // 9 = A4. Need to add enum for potential values, but default to A4 for now
        public int PaperSize { get { return _paperSize; } private set { _paperSize = value; } }
        private int _scale = 100;
        public int Scale { get { return _scale; } set { if (value > 400) _scale = 400; else _scale = value; } }
        public bool UseFirstPageNumber { get; set; }
        public bool UsePrinterDefaults { get; set; }


        /***********************************
         * CONSTRUCTORS
         ************************************/


        public PageSetup(Worksheet worksheet)
        {
            Worksheet = worksheet;
        }

        public PageSetup(Worksheet worksheet,
            bool blackAndWhite, int copies, bool draft, int firstPageNumber, int fitToHeight, int fitToWidth, int horizontalDpi, int verticalDpi,
            string relationshipId, WorksheetOrientation orientation, int paperSize, int scale, bool useFirstPageNumber, bool usePrinterDefaults)
            : this (worksheet)
        {
            BlackAndWhite = blackAndWhite;
            Copies = copies;
            Draft = draft;
            FirstPageNumber = firstPageNumber;
            FitToHeight = fitToHeight;
            FitToWidth = fitToWidth;
            HorizontalDpi = horizontalDpi;
            VerticalDpi = verticalDpi;
            RelationshipId = relationshipId;
            Orientation = orientation;
            PaperSize = paperSize;
            Scale = scale;
            UseFirstPageNumber = useFirstPageNumber;
            UsePrinterDefaults = usePrinterDefaults;
        }


        /***********************************
         * INTERNAL METHODS
         ************************************/


        internal PageSetup Clone(Worksheet worksheet)
        {
            PageSetup newPageSetup = new PageSetup(worksheet,
                BlackAndWhite, Copies, Draft, FirstPageNumber, FitToHeight, FitToWidth, HorizontalDpi, VerticalDpi, RelationshipId, Orientation, 
                PaperSize, Scale, UseFirstPageNumber, UsePrinterDefaults);
            return newPageSetup;
        }

        internal bool HasValue()
        {
            return (BlackAndWhite ||
                Copies > 1 ||
                Draft ||
                FirstPageNumber > 1 ||
                FitToHeight != -1 ||
                FitToWidth != -1 ||
                HorizontalDpi != 600 ||
                VerticalDpi != 600 ||
                RelationshipId != "" ||
                Orientation != WorksheetOrientation.Portrait ||
                PaperSize != 9 ||
                Scale != 400 ||
                UseFirstPageNumber ||
                UsePrinterDefaults);
        }

        /***********************************
         * DAL METHODS
         ************************************/

        // Read

        internal static PageSetup ReadPageSetupFromReader(CustomOpenXmlReader reader, Worksheet worksheet)
        {
            PageSetup pageSetup = new PageSetup(worksheet);

            foreach (CustomOpenXmlAttribute attribute in reader.Attributes)
            {
                switch (attribute.LocalName)
                {
                    case "blackAndWhite":
                        pageSetup.BlackAndWhite = attribute.GetBoolValue();
                        break;
                    case "copies":
                        pageSetup.Copies = attribute.GetIntValue();
                        break;
                    case "draft":
                        pageSetup.Draft = attribute.GetBoolValue();
                        break;
                    case "firstPageNumber":
                        pageSetup.FirstPageNumber = attribute.GetIntValue();
                        break;
                    case "fitToHeight":
                        pageSetup.FitToHeight = attribute.GetIntValue();
                        break;
                    case "fitToWidth":
                        pageSetup.FitToWidth = attribute.GetIntValue();
                        break;
                    case "horizontalDpi":
                        pageSetup.HorizontalDpi = attribute.GetIntValue();
                        break;
                    case "verticalDpi":
                        pageSetup.VerticalDpi = attribute.GetIntValue();
                        break;
                    case "id":
                        pageSetup.RelationshipId = attribute.Value;
                        break;
                    case "orientation":
                        switch (attribute.Value)
                        {
                            case "landscape":
                                pageSetup.Orientation = WorksheetOrientation.Landscape;
                                break;
                            case "portrait":
                                pageSetup.Orientation = WorksheetOrientation.Portrait;
                                break;
                        }
                        break;
                    case "paperSize":
                        pageSetup.PaperSize = attribute.GetIntValue();
                        break;
                    case "scale":
                        pageSetup.Scale = attribute.GetIntValue();
                        break;
                    case "useFirstPageNumber":
                        pageSetup.UseFirstPageNumber = attribute.GetBoolValue();
                        break;
                    case "usePrinterDefaults":
                        pageSetup.UsePrinterDefaults = attribute.GetBoolValue();
                        break;
                    default:
                        throw new Exception(string.Format("PageSetup attribute {0} not coded", attribute.LocalName));
                }
            }

            return pageSetup;
        }

        // Write

        internal static void WritePageSetupToWriter(CustomOpenXmlWriter<OpenXmlPackaging.WorksheetPart> writer, PageSetup pageSetup)
        {
            if (pageSetup.HasValue())
            {
                writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.PageSetup());

                if (pageSetup.BlackAndWhite) writer.WriteAttribute("blackAndWhite", pageSetup.BlackAndWhite);
                if (pageSetup.Copies > 1) writer.WriteAttribute("copies", pageSetup.Copies);
                if (pageSetup.Draft) writer.WriteAttribute("draft", pageSetup.Draft);
                if (pageSetup.FirstPageNumber > 1) writer.WriteAttribute("firstPageNumber", pageSetup.FirstPageNumber);
                if (pageSetup.FitToHeight > -1) writer.WriteAttribute("fitToHeight", pageSetup.FitToHeight);
                if (pageSetup.FitToWidth > -1) writer.WriteAttribute("fitToWidth", pageSetup.FitToWidth);
                if (pageSetup.HorizontalDpi != 600) writer.WriteAttribute("horizontalDpi", pageSetup.HorizontalDpi);
                if (pageSetup.VerticalDpi != 600) writer.WriteAttribute("verticalDpi", pageSetup.VerticalDpi);
                if (pageSetup.Orientation != WorksheetOrientation.Portrait) writer.WriteAttribute("orientation", Helpers.ToCamelCase(pageSetup.Orientation.ToString()));
                if (pageSetup.PaperSize != 9) writer.WriteAttribute("paperSize", pageSetup.PaperSize);
                if (pageSetup.RelationshipId != "") writer.WriteAttribute("id", pageSetup.RelationshipId, "r");
                if (pageSetup.Scale != 400) writer.WriteAttribute("scale", pageSetup.Scale);
                if (pageSetup.UseFirstPageNumber) writer.WriteAttribute("useFirstPageNumber", pageSetup.UseFirstPageNumber);
                if (pageSetup.UsePrinterDefaults) writer.WriteAttribute("usePrinterDefaults", pageSetup.UsePrinterDefaults);

                writer.WriteEndElement();   // PageSetup
            }
        }
    }
}
