using System;
using System.Collections.Generic;
using System.Linq;

using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{
    // This class has no meaning in the business object model. 
    // It is merely a convenient class to tie all the data from the same worksheet xml file together.
    // If any one of the data items here is saved they all need re-saving to the xml file, so need to be read anyway.
    // It could be optimised to lazy load if the workbook is read only.

    internal class WorksheetData : BaseExcel
    {
        public Worksheet Worksheet { get; private set; }

        private SheetViews _sheetViews;
        public SheetViews SheetViews { get { if (_sheetViews == null) _sheetViews = new SheetViews(Worksheet); return _sheetViews; } private set { _sheetViews = value; } }

        private WorksheetFormat _format;
        public WorksheetFormat Format { get { if (_format == null) _format = new WorksheetFormat(Worksheet); return _format; } private set { _format = value; } }

        private ColumnCollection _columns;
        public ColumnCollection Columns { get { if (_columns == null) _columns = new ColumnCollection(Worksheet); return _columns; } private set { _columns = value; } }

        private RowCollection _rows;
        public RowCollection Rows { get { if (_rows == null) _rows = new RowCollection(Worksheet); return _rows; } private set { _rows = value; } }        

        private PrintOptions _printOptions;
        public PrintOptions PrintOptions { get { if (_printOptions == null) _printOptions = new PrintOptions(Worksheet); return _printOptions; } private set { _printOptions = value; } }

        private PageMargins _pageMargins;
        public PageMargins PageMargins { get { if (_pageMargins == null) _pageMargins = new PageMargins(Worksheet); return _pageMargins; } private set { _pageMargins = value; } }

        private PageSetup _pageSetup;
        public PageSetup PageSetup { get { if (_pageSetup == null) _pageSetup = new PageSetup(Worksheet); return _pageSetup; } private set { _pageSetup = value; } }

        private HeaderFooter _headerFooter;
        public HeaderFooter HeaderFooter { get { if (_headerFooter == null) _headerFooter = new HeaderFooter(Worksheet); return _headerFooter; } private set { _headerFooter = value; } }

        // temporary implementation
        private Drawing _drawing;
        public Drawing Drawing { get { return _drawing; } private set { _drawing = value; } }
        
        /***********************************
         * CONSTRUCTORS
         ************************************/

        public WorksheetData(Worksheet worksheet)
        {
            Worksheet = worksheet;
        }


        /***********************************
         * PUBLIC PROPERTIES
         ************************************/

        /***********************************
         * PUBLIC METHODS
         ************************************/

        public void CloneToWorksheet(Worksheet newWorksheet)
        {
            // This method needs a better name, and perhaps a more orthodox way of setting the worksheet data.
            newWorksheet.WorksheetData.SheetViews = SheetViews.Clone(newWorksheet);
            newWorksheet.WorksheetData.Columns = Columns.Clone(newWorksheet);
            newWorksheet.WorksheetData.Rows = Rows.Clone(newWorksheet);
            newWorksheet.WorksheetData.Format = Format.Clone(newWorksheet);
            newWorksheet.WorksheetData.PrintOptions = PrintOptions.Clone(newWorksheet);
            newWorksheet.WorksheetData.PageMargins = PageMargins.Clone(newWorksheet);
            newWorksheet.WorksheetData.PageSetup = PageSetup.Clone(newWorksheet);
            newWorksheet.WorksheetData.HeaderFooter = HeaderFooter.Clone(newWorksheet);

            if (Worksheet.PrintArea != "")
                newWorksheet.PrintArea = Worksheet.PrintArea;
            if (Worksheet.PrintTitles != "")
                newWorksheet.PrintTitles = Worksheet.PrintTitles;

            if (Drawing != null)
                newWorksheet.WorksheetData.Drawing = Drawing.Clone(newWorksheet);

            // create relationships.            
            newWorksheet.PageSetup.RelationshipId = "";

            OpenXmlPackaging.WorksheetPart newWorksheetPart = Worksheet.GetWorksheetPartByWorksheet(newWorksheet);
            OpenXmlPackaging.WorksheetPart worksheetPart = Worksheet.GetWorksheetPartByWorksheet(Worksheet);
                
            IEnumerable<OpenXmlPackaging.SpreadsheetPrinterSettingsPart> printerSettingsParts = worksheetPart.SpreadsheetPrinterSettingsParts;
            foreach (OpenXmlPackaging.SpreadsheetPrinterSettingsPart printerSettingsPart in printerSettingsParts)
            {                        
                string relationshipId = newWorksheetPart.CreateRelationshipToPart(printerSettingsPart);
                newWorksheet.PageSetup.RelationshipId = relationshipId;
            }            
        }

        /***********************************
         * INTERNAL METHODS
         ************************************/

        internal void Save()
        {
            WorksheetData.SaveWorksheetData(this);
        }

        
        internal void LoadWorksheetData()
        {            
            // Ideally this would be called privately instead of internally, but is good to get it working.
            WorksheetData.ReadWorksheetDataFromWorksheetData(this);
        }


        /***********************************
         * DAL METHODS
         ************************************/


        // Read
        
        private static void ReadWorksheetDataFromWorksheetData(WorksheetData worksheetData)
        {
            OpenXmlSpreadsheet.Worksheet worksheetElement = Worksheet.GetWorksheetElementFromWorksheet(worksheetData.Worksheet);

            using (CustomOpenXmlReader reader = CustomOpenXmlReader.Create(worksheetElement))
            {
                while (reader.Read())
                {
                    if (reader.IsStartElementOfType<OpenXmlSpreadsheet.SheetViews>())
                        worksheetData.SheetViews = SheetViews.ReadSheetViewsFromReader(reader, worksheetData.Worksheet);
                    else if (reader.IsStartElementOfType<OpenXmlSpreadsheet.SheetFormatProperties>())
                        worksheetData.Format = WorksheetFormat.ReadWorksheetFormatFromReader(reader, worksheetData.Worksheet);
                    else if (reader.IsStartElementOfType<OpenXmlSpreadsheet.Columns>())
                        worksheetData.Columns = ColumnCollection.ReadColumnsFromReader(reader, worksheetData.Worksheet);
                    else if (reader.IsStartElementOfType<OpenXmlSpreadsheet.SheetData>())
                        worksheetData.Rows = RowCollection.ReadRowsFromReader(reader, worksheetData.Worksheet);
                    else if (reader.IsStartElementOfType<OpenXmlSpreadsheet.MergeCells>())
                        MergeCell.UpdateCellMergeCellsFromReader(reader, worksheetData.Worksheet);
                    else if (reader.IsStartElementOfType<OpenXmlSpreadsheet.Hyperlinks>())
                        Hyperlink.UpdateCellHyperlinksFromReader(reader, worksheetData.Worksheet);
                    else if (reader.IsStartElementOfType<OpenXmlSpreadsheet.PrintOptions>())
                        worksheetData.PrintOptions = PrintOptions.ReadPrintOptionsFromReader(reader, worksheetData.Worksheet);
                    else if (reader.IsStartElementOfType<OpenXmlSpreadsheet.PageMargins>())
                        worksheetData.PageMargins = PageMargins.ReadPageMarginsFromReader(reader, worksheetData.Worksheet);
                    else if (reader.IsStartElementOfType<OpenXmlSpreadsheet.PageSetup>())
                        worksheetData.PageSetup = PageSetup.ReadPageSetupFromReader(reader, worksheetData.Worksheet);
                    else if (reader.IsStartElementOfType<OpenXmlSpreadsheet.HeaderFooter>())
                        worksheetData.HeaderFooter = HeaderFooter.ReadHeaderFooterFromReader(reader, worksheetData.Worksheet);
                    else if (reader.IsStartElementOfType<OpenXmlSpreadsheet.Drawing>())
                        worksheetData.Drawing = Drawing.ReadDrawingFromReader(reader, worksheetData.Worksheet);
                }
            }
        }

        // Write

        private static void SaveWorksheetData(WorksheetData worksheetData)
        {
            OpenXmlPackaging.WorksheetPart worksheetPart = Worksheet.GetWorksheetPartByWorksheet(worksheetData.Worksheet);
            
            using (CustomOpenXmlWriter<OpenXmlPackaging.WorksheetPart> writer = new CustomOpenXmlWriter<OpenXmlPackaging.WorksheetPart>(worksheetPart))
            {
                writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.Worksheet());

                if (worksheetData._rows != null)
                    RowCollection.WriteSheetDimensionToWriter(writer, worksheetData.Worksheet);
                if (worksheetData._sheetViews != null)
                    SheetViews.WriteSheetViewsToWriter(writer, worksheetData.SheetViews);
                if (worksheetData._format != null)
                    WorksheetFormat.WriteWorksheetFormatToWriter(writer, worksheetData.Format);
                if (worksheetData._columns != null)
                    ColumnCollection.WriteColumnsToWriter(writer, worksheetData.Columns);
                if (worksheetData._rows != null)
                    RowCollection.WriteRowsToWriter(writer, worksheetData.Rows);               
                // should probably store merge cells and hyperlinks in their own objects. Reference them from the cells
                MergeCell.WriteMergeCellsToWriter(writer, worksheetData.Worksheet);
                Hyperlink.WriteHyperlinksToWriter(writer, worksheetData.Worksheet);
                if (worksheetData._printOptions != null)
                    PrintOptions.WritePrintOptionsToWriter(writer, worksheetData.PrintOptions);
                if (worksheetData._pageMargins != null)
                    PageMargins.WritePageMarginsToWriter(writer, worksheetData.PageMargins);
                if (worksheetData._pageSetup != null)
                    PageSetup.WritePageSetupToWriter(writer, worksheetData.PageSetup);
                if (worksheetData._headerFooter != null)
                    HeaderFooter.WriteHeaderFooterToWriter(writer, worksheetData.HeaderFooter);
                if (worksheetData._drawing != null)
                    Drawing.WriteDrawingToWorksheetWriter(writer, worksheetData.Drawing);                    

                writer.WriteEndElement();   // Worksheet
            }

            if (worksheetData._drawing != null)
                worksheetData.Drawing.Save();
        }
    }
}
