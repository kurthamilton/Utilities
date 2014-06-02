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
    internal class BordersCollection : BaseExcel, IEnumerable<Borders>
    {
        private Styles Styles { get; set; }

        private SortedDictionary<int, Borders> _bordersDictionary;
        private SortedDictionary<int, Borders> BordersDictionary { get { if (_bordersDictionary == null) _bordersDictionary = new SortedDictionary<int,Borders>(); return _bordersDictionary; } }
        
        public int Count { get { return BordersDictionary.Count; } }

        /***********************************
        * CONSTRUCTORS
        ************************************/

        public BordersCollection(Styles styles)
        {
            Styles = styles;
        }        

        /***********************************
        * PUBLIC PROPERTIES
        ************************************/
        public Borders DefaultBorders
        {
            get
            {
                if (BordersDictionary.Count > 0)
                    return BordersDictionary[Borders.DefaultBordersId];
                else
                    return null;
            }
        }

        /***********************************
        * PUBLIC METHODS
        ************************************/
        public Borders this[int index]
        {
            get
            {
                if (Contains(index))
                    return BordersDictionary[index];
                else
                    return null;
            }
        }

        public Borders this[Borders borders]
        {
            get
            {
                if (borders != null)
                    return (BordersDictionary.Values.FirstOrDefault(b => b.Equals(borders)));
                else
                    return null;
            }
        }

        public void Clear()
        {
            // first borders is default borders. Don't clear.
            while (BordersDictionary.Count > 1)
            {
                BordersDictionary.Remove(BordersDictionary.Max(b => b.Key));
            }
        }

        public bool Contains(int bordersId)
        {
            return BordersDictionary.ContainsKey(bordersId);
        }

        public Borders Insert(Borders borders)
        {
            Borders existingBorders = this[borders];

            if (existingBorders == null)
            {
                int newBordersId = GenerateNewBordersId();
                Borders newBorders = new Borders(borders.Styles, newBordersId, 
                    borders.Left.Clone(), borders.Right.Clone(), borders.Top.Clone(), borders.Bottom.Clone(), borders.Diagonal.Clone());
                AddBordersToCollection(newBorders);
                return newBorders;
            }

            return existingBorders;
        }

        public void Delete(int index)
        {
            if (Contains(index))
                BordersDictionary.Remove(index);
        }

        // Implement IEnumerable
        public IEnumerator<Borders> GetEnumerator()
        {
            return new GenericEnumerator<Borders>(BordersDictionary);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }


        /***********************************
        * PRIVATE METHODS
        ************************************/
        private int GenerateNewBordersId()
        {
            return OfficeUtilities.GetFirstUnusedKeyFromCollection<Borders>(BordersDictionary);
        }

        private void AddBordersToCollection(Borders borders)
        {
            if (!Contains(borders.BordersId))
                BordersDictionary.Add(borders.BordersId, borders);
        }

        /***********************************
        * DAL METHODS
        ************************************/

        // Read

        internal static BordersCollection ReadBordersCollectionFromReader(CustomOpenXmlReader reader, Styles styles)
        {
            BordersCollection bordersCollection = new BordersCollection(styles);

            while (reader.ReadToEndElement<OpenXmlSpreadsheet.Borders>())
            {
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.Border>())
                {
                    Borders borders = Borders.ReadBordersFromReader(reader, styles, bordersCollection.Count + Borders.DefaultBordersId);
                    bordersCollection.AddBordersToCollection(borders);
                }
            }

            return bordersCollection;
        }

        // Write

        internal static void WriteBordersToWriter(CustomOpenXmlWriter<OpenXmlPackaging.WorkbookStylesPart> writer, IEnumerable<Borders> bordersCollection)
        {
            writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.Borders());

            foreach (Borders borders in bordersCollection)
            {
                Borders.WriteBordersToWriter(writer, borders);
            }

            writer.WriteEndElement();   // Borders
        }
    }
}
