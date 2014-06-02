using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{
    internal class SharedStringCollection : BaseExcel, IEnumerable<SharedString>
    {
        public Workbook Workbook { get; private set; }

        private SortedDictionary<int, SharedString> SharedStrings { get; set; }
        private Dictionary<string, int> StringDictionary { get; set; }

        public int Count 
        { 
            get 
            {                 
                return SharedStrings.Count; 
            } 
        }

        /***********************************
         * CONSTRUCTORS
         ************************************/   

        public SharedStringCollection(Workbook workbook)
        {
            SharedStrings = new SortedDictionary<int, SharedString>();
            StringDictionary = new Dictionary<string, int>();
            Workbook = workbook;
            GetExistingSharedStrings();
        }


        /***********************************
         * public METHODS
         ************************************/

        // Implement IEnumerable
        public IEnumerator<SharedString> GetEnumerator()
        {
            return new GenericEnumerator<SharedString>(SharedStrings);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        /***********************************
         * INTERNAL METHODS
         ************************************/


        internal SharedString this[int index]
        {
            get
            {
                if (SharedStrings.ContainsKey(index))
                    return SharedStrings[index];
                return null;
            }
        }

        internal int this[string sharedString]
        {
            get
            {
                Add(sharedString);
                return StringDictionary[sharedString];
            }
        }

        internal int this[SharedString sharedString]
        {
            get
            {
                Add(sharedString);
                return sharedString.Index;
            }
        }

        internal void Save()
        {
            SharedStringCollection.SaveSharedStrings(this);
        }

        /***********************************
         * PRIVATE METHODS
         ************************************/


        private void GetExistingSharedStrings()
        {            
            IEnumerable<SharedString> sharedStrings = SharedStringCollection.GetSharedStringsFromWorkbook(Workbook);
            foreach (SharedString sharedString in sharedStrings)
            {
                AddDictionaryEntry(sharedString);
            }
        }

        private void Add(string sharedStringValue)
        {
            if (!Contains(sharedStringValue))
            {
                SharedString sharedString = new SharedString(Workbook, sharedStringValue);
                AddDictionaryEntry(sharedString);
            }
        }
        private void Add(SharedString sharedString)
        {
            if (!Contains(sharedString))
                AddDictionaryEntry(sharedString);
            if (sharedString.Index < 0)
                sharedString.Index = SharedStrings.Values.First(s => s.IsFontString && s.Equals(sharedString)).Index;
        }

        private void AddDictionaryEntry(SharedString sharedString)
        {
            sharedString.Index = SharedStrings.Count;

            SharedStrings.Add(sharedString.Index, sharedString);
            // only store simple strings in the StringDictionary
            if (!sharedString.IsFontString)
                StringDictionary.Add(sharedString.ToString(), sharedString.Index);
        }

        
        private bool Contains(SharedString sharedString)
        {
            if (sharedString.IsFontString)
                return (SharedStrings.Values.FirstOrDefault(s => s.IsFontString && s.Equals(sharedString)) != null);
            else
                return Contains(sharedString.ToString());
        }
        private bool Contains(string sharedString)
        {
            return StringDictionary.ContainsKey(sharedString);
        }

        /***********************************
         * DAL METHODS
         ************************************/

        // Read
        internal static IEnumerable<SharedString> GetSharedStringsFromWorkbook(Workbook workbook)
        {
            OpenXmlPackaging.SharedStringTablePart sharedStringTablePart = workbook.Document.WorkbookPart.SharedStringTablePart;
            if (sharedStringTablePart != null)
            {
                OpenXmlSpreadsheet.SharedStringTable sharedStringTable = sharedStringTablePart.SharedStringTable;

                IEnumerable<SharedString> sharedStrings;

                using (CustomOpenXmlReader reader = CustomOpenXmlReader.Create(sharedStringTable))
                {
                    sharedStrings = ReadSharedStringsFromReader(reader, workbook);
                }

                return sharedStrings;
            }
            else
            {
                return new List<SharedString>();
            }
        }

        private static IEnumerable<SharedString> ReadSharedStringsFromReader(CustomOpenXmlReader reader, Workbook workbook)
        {
            List<SharedString> sharedStrings = new List<SharedString>();

            while (reader.ReadToEndElement<OpenXmlSpreadsheet.SharedStringTable>())
            {
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.SharedStringItem>())
                {
                    SharedString sharedString = SharedString.ReadSharedStringFromReader(reader, workbook);
                    sharedStrings.Add(sharedString);
                }
            }

            return sharedStrings;
        }

        // Write
        internal static void SaveSharedStrings(SharedStringCollection sharedStrings)
        {
            using (CustomOpenXmlWriter<OpenXmlPackaging.SharedStringTablePart> writer =
                new CustomOpenXmlWriter<OpenXmlPackaging.SharedStringTablePart>(sharedStrings.Workbook.Document.WorkbookPart.SharedStringTablePart)
            )
            {
                SharedStringCollection.WriteSharedStringsToWriter(writer, sharedStrings);
            }
        }

        internal static void WriteSharedStringsToWriter(CustomOpenXmlWriter<OpenXmlPackaging.SharedStringTablePart> writer, IEnumerable<SharedString> sharedStrings)
        {
            writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.SharedStringTable());

            foreach (SharedString sharedString in sharedStrings)
            {
                SharedString.WriteSharedStringToWriter(writer, sharedString);
            }

            writer.WriteEndElement();   // SharedStringTable
        }
    }
}
