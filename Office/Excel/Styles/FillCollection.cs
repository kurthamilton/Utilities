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
    internal class FillCollection : BaseExcel, IEnumerable<Fill>
    {
        private Styles Styles { get; set; }

        private SortedDictionary<int, Fill> _fillDictionary;
        private SortedDictionary<int, Fill> FillDictionary { get { if (_fillDictionary == null) _fillDictionary = new SortedDictionary<int,Fill>(); return _fillDictionary; } }
        
        public int Count { get { return FillDictionary.Count; } }

        /***********************************
        * CONSTRUCTORS
        ************************************/

        public FillCollection(Styles styles)
        {
            Styles = styles;
        }

        /***********************************
        * PUBLIC PROPERTIES
        ************************************/
        public Fill DefaultFill
        {
            get
            {
                if (FillDictionary.Count > 0)
                    return FillDictionary[Fill.DefaultFillId];
                else
                    return null;
            }
        }

        /***********************************
        * PUBLIC METHODS
        ************************************/
        public Fill this[int index]
        {
            get
            {
                if (Contains(index))
                    return FillDictionary[index];
                else
                    return null;
            }
        }

        public Fill this[Fill fill]
        {
            get
            {
                if (fill != null)
                    return FillDictionary.Values.ToList().Find(f => f.Equals(fill));
                else
                    return null;
            }
        }

        public void Clear()
        {
            // first fill is default fill. Don't clear.
            while (FillDictionary.Count > 1)
            {
                FillDictionary.Remove(FillDictionary.Max(f => f.Key));
            }
        }

        public bool Contains(int fillId)
        {
            return FillDictionary.ContainsKey(fillId);
        }

        public Fill Insert(Fill fill)
        {
            Fill existingFill = this[fill];

            if (existingFill == null)
            {
                int newFillId = GenerateNewFillId();
                Fill newFill = new Fill(fill.Styles, newFillId, fill.PatternType, fill.ForegroundColor, fill.BackgroundColor);
                AddFillToCollection(newFill);
                return newFill;
            }

            return existingFill;
        }

        public void Delete(int index)
        {
            if (Contains(index))
                FillDictionary.Remove(index);
        }

        // Implement IEnumerable
        public IEnumerator<Fill> GetEnumerator()
        {
            return new GenericEnumerator<Fill>(FillDictionary);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        /***********************************
        * PRIVATE METHODS
        ************************************/
        private int GenerateNewFillId()
        {
            return OfficeUtilities.GetFirstUnusedKeyFromCollection<Fill>(FillDictionary);
        }

        private void AddFillToCollection(Fill fill)
        {
            if (!Contains(fill.FillId))
                FillDictionary.Add(fill.FillId, fill);
        }

        /***********************************
        * DAL METHODS
        ************************************/

        // Read

        internal static FillCollection ReadFillsFromReader(CustomOpenXmlReader reader, Styles styles)
        {
            FillCollection fills = new FillCollection(styles);

            while (reader.ReadToEndElement<OpenXmlSpreadsheet.Fills>())
            {
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.Fill>())
                {
                    Fill fill = Fill.ReadFillFromReader(reader, styles, fills.Count + Fill.DefaultFillId);
                    fills.AddFillToCollection(fill);
                }
            }

            return fills;
        }

        // Write

        internal static void WriteFillsToWriter(CustomOpenXmlWriter<OpenXmlPackaging.WorkbookStylesPart> writer, IEnumerable<Fill> fills)
        {
            writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.Fills());

            foreach (Fill fill in fills)
            {
                Fill.WriteFillToWriter(writer, fill);
            }

            writer.WriteEndElement();   // Fills
        }
    }
}
