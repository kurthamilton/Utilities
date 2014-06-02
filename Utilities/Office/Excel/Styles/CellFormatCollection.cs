using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using System.Xml;
using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{
    internal class CellFormatCollection : BaseExcel, IEnumerable<CellFormat>
    {
        private Styles Styles { get; set; }

        private SortedDictionary<int, CellFormat> _cellFormatDictionary;
        private SortedDictionary<int, CellFormat> CellFormatDictionary { get { if (_cellFormatDictionary == null) _cellFormatDictionary = new SortedDictionary<int,CellFormat>(); return _cellFormatDictionary; } set { _cellFormatDictionary = value; } }
        
        public int Count { get { return CellFormatDictionary.Count; } }

        /***********************************
         * CONSTRUCTORS
         ************************************/

        public CellFormatCollection(Styles styles)
        {
            Styles = styles;
        }        


        /***********************************
         * PUBLIC METHODS
         ************************************/

        public CellFormat this[int index]
        {
            get
            {
                if (Contains(index))
                    return CellFormatDictionary[index];
                else
                    return null;
            }
        }
        public CellFormat this[CellFormat cellFormat]
        {
            get
            {
                if (cellFormat != null)
                    return (CellFormatDictionary.Values.FirstOrDefault(f => f.Equals(cellFormat)));
                else
                    return null;
            }
        }

        public void Clear()
        {
            // first cellFormat is default font. Don't clear.
            while (CellFormatDictionary.Count > 1)
            {
                CellFormatDictionary.Remove(CellFormatDictionary.Max(f => f.Key));
            }
        }

        public bool Contains(int cellFormatId)
        {
            return CellFormatDictionary.ContainsKey(cellFormatId);
        }

        public CellFormat Insert(CellFormat cellFormat)
        {
            CellFormat existingCellFormat = this[cellFormat];

            if (existingCellFormat == null)
            {
                int newCellFormatId = GenerateNewCellFormatId();
                CellFormat newCellFormat = 
                    new CellFormat(cellFormat.Styles, cellFormat.BaseFormatId, newCellFormatId, cellFormat.Alignment, 
                        cellFormat.Borders, cellFormat.Fill, cellFormat.Font, cellFormat.NumberFormat);
                AddCellFormatToCollection(newCellFormat);
                return newCellFormat;
            }

            return existingCellFormat;
        }

        public void Delete(int index)
        {
            if (Contains(index))
                CellFormatDictionary.Remove(index);
        }

        // Implement IEnumerable
        public IEnumerator<CellFormat> GetEnumerator()
        {
            return new GenericEnumerator<CellFormat>(CellFormatDictionary);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }


        /***********************************
         * PRIVATE METHODS
         ************************************/

        private void AddCellFormatToCollection(CellFormat cellFormat)
        {
            if (!Contains(cellFormat.CellFormatId))
                CellFormatDictionary.Add(cellFormat.CellFormatId, cellFormat);
        }

        private int GenerateNewCellFormatId()
        {
            return OfficeUtilities.GetFirstUnusedKeyFromCollection<CellFormat>(CellFormatDictionary);
        }        

        /***********************************
         * DAL METHODS
         ************************************/

        // Read

        internal static CellFormatCollection ReadCellFormatsFromReader(CustomOpenXmlReader reader, Styles styles)
        {
            CellFormatCollection cellFormats = new CellFormatCollection(styles);
            
            while (reader.ReadToEndElement<OpenXmlSpreadsheet.CellFormats>())
            {
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.CellFormat>())
                {
                    CellFormat cellFormat = CellFormat.ReadCellFormatFromReader(reader, styles, cellFormats.Count + CellFormat.DefaultStyleIndex);
                    cellFormats.AddCellFormatToCollection(cellFormat);
                }
            }

            return cellFormats;
        }

        // Write

        internal static void WriteCellFormatsToWriter(CustomOpenXmlWriter<OpenXmlPackaging.WorkbookStylesPart> writer, IEnumerable<CellFormat> cellFormats)
        {
            writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.CellFormats());

            foreach (CellFormat cellFormat in cellFormats)
            {
                CellFormat.WriteCellFormatToWriter(writer, cellFormat);
            }

            writer.WriteEndElement();   // CellFormats
        }
    }
}
