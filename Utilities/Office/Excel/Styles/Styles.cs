using System;
using System.Collections.Generic;

using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{
    // keep internal for now until fully developed
    internal class Styles : BaseExcel
    {
        public Workbook Workbook { get; private set; }

        private NumberFormatCollection _numberFormats;
        public NumberFormatCollection NumberFormats { get { if (_numberFormats == null) _numberFormats = new NumberFormatCollection(this); return _numberFormats; } private set { _numberFormats = value; } }

        private FontCollection _fonts;
        internal FontCollection Fonts { get { if (_fonts == null) _fonts = new FontCollection(this); return _fonts; } private set { _fonts = value; } }

        private FillCollection _fills;
        internal FillCollection Fills { get { if (_fills == null) _fills = new FillCollection(this); return _fills; } private set { _fills = value; } }        

        private BordersCollection _borders;
        internal BordersCollection Borders { get { if (_borders == null) _borders = new BordersCollection(this); return _borders; } private set { _borders = value; } }
        
        private CellFormatCollection _cellFormats;
        internal CellFormatCollection CellFormats { get { if (_cellFormats == null) _cellFormats = new CellFormatCollection(this); return _cellFormats; } private set { _cellFormats = value; } }

        private List<string> IndexedColors { get; set; }

        /***********************************
         * CONSTRUCTORS
         ************************************/


        public Styles(Workbook workbook)
        {            
            Workbook = workbook;
            LoadStyles();
        }


        /***********************************
        * INTERNAL METHODS
        ************************************/

        internal void Save()
        {
            Styles.SaveStyles(this);
        }

        /***********************************
        * PRIVATE METHODS
        ************************************/

        private void LoadStyles()
        {
            Styles.ReadStylesFromStyles(this);
        }

        /***********************************
        * DAL METHODS
        ************************************/     

        // Read

        private static void ReadStylesFromStyles(Styles styles)
        {
            OpenXmlSpreadsheet.Stylesheet stylesheet = styles.Workbook.Document.WorkbookPart.WorkbookStylesPart.Stylesheet;
            
            using (CustomOpenXmlReader reader = CustomOpenXmlReader.Create(stylesheet))
            {
                while (reader.ReadToEndElement<OpenXmlSpreadsheet.Stylesheet>())
                {                    
                    if (reader.IsStartElementOfType<OpenXmlSpreadsheet.NumberingFormats>())
                        styles.NumberFormats = NumberFormatCollection.ReadNumberFormatsFromReader(reader, styles);
                    else if (reader.IsStartElementOfType<OpenXmlSpreadsheet.Fonts>())
                        styles.Fonts = FontCollection.ReadFontsFromReader(reader, styles);
                    else if (reader.IsStartElementOfType<OpenXmlSpreadsheet.Fills>())
                        styles.Fills = FillCollection.ReadFillsFromReader(reader, styles);
                    else if (reader.IsStartElementOfType<OpenXmlSpreadsheet.Borders>())
                        styles.Borders = BordersCollection.ReadBordersCollectionFromReader(reader, styles);
                    // not sure if required.
                    //else if (reader.IsStartElementOfType<OpenXmlSpreadsheet.CellStyleFormats>())
                        //cellStyleFormats = CellFormatCollection.ReadCellStyleFormatsFromReader(reader, styles);
                    else if (reader.IsStartElementOfType<OpenXmlSpreadsheet.CellFormats>())
                        styles.CellFormats = CellFormatCollection.ReadCellFormatsFromReader(reader, styles);
                    else if (reader.IsStartElementOfType<OpenXmlSpreadsheet.Colors>())
                    {
                        while (reader.ReadToEndElement<OpenXmlSpreadsheet.Colors>())
                        {
                            if (reader.IsStartElementOfType<OpenXmlSpreadsheet.IndexedColors>())
                                styles.IndexedColors = Color.ReadIndexedColorsFromReader(reader);
                        }
                    }
                }
            }
        }
    

        // Write

        private static void SaveStyles(Styles styles)
        {
            using (CustomOpenXmlWriter<OpenXmlPackaging.WorkbookStylesPart> writer =
                new CustomOpenXmlWriter<OpenXmlPackaging.WorkbookStylesPart>(styles.Workbook.Document.WorkbookPart.WorkbookStylesPart)
            )
            {
                Styles.WriteStylesToWriter(writer, styles);
            }
        }

        private static void WriteStylesToWriter(CustomOpenXmlWriter<OpenXmlPackaging.WorkbookStylesPart> writer, Styles styles)
        {
            writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.Stylesheet());

            if (styles._numberFormats != null) NumberFormatCollection.WriteNumberFormatsToWriter(writer, styles.NumberFormats);
            if (styles._fonts != null) FontCollection.WriteFontsToWriter(writer, styles.Fonts);
            if (styles._fills != null) FillCollection.WriteFillsToWriter(writer, styles.Fills);
            if (styles._borders != null) BordersCollection.WriteBordersToWriter(writer, styles.Borders);
            if (styles._cellFormats != null) CellFormatCollection.WriteCellFormatsToWriter(writer, styles.CellFormats);
            if (styles.IndexedColors != null) Color.WriteIndexedColorsToWriter(writer, styles.IndexedColors);

            writer.WriteEndElement();   // Stylesheet
        }
    }
}
