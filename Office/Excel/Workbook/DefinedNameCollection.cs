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
    internal class DefinedNameCollection : BaseExcel, IEnumerable<DefinedName>
    {
        private Workbook Workbook { get; set; }

        private SortedDictionary<int, DefinedName> NamesDictionary { get; set; }

        public int Count { get { return NamesDictionary.Count; } }

        private const string PrintTitlesName = "_xlnm.Print_Titles";
        private const string PrintAreaName = "_xlnm.Print_Area";

        /***********************************
         * CONSTRUCTORS
         ************************************/

        public DefinedNameCollection(Workbook workbook)
        {
            Workbook = workbook;
            NamesDictionary = new SortedDictionary<int, DefinedName>();
        }
               


        /***********************************
         * PUBLIC METHODS
         ************************************/


        public DefinedName this[Worksheet worksheet, string name, DefinedNameScope scope]
        {
            get
            {
                return NamesDictionary.Values
                    .FirstOrDefault(n => 
                        n.Worksheet.Equals(worksheet) && 
                        string.Compare(n.Name, name, StringComparison.InvariantCultureIgnoreCase) == 0 && 
                        n.Scope == scope);
            }
        }

        public bool Contains(int nameId)
        {
            return NamesDictionary.ContainsKey(nameId);
        }

        public DefinedName Insert(DefinedName definedName)
        {
            DefinedName existingName = this[definedName.Worksheet, definedName.Name, definedName.Scope];

            if (existingName != null)
            {
                existingName.SetAddresses(definedName.GetAddressString(false));
                return existingName;
            }
            else
            {
                int nameId = GenerateNewNameId();

                DefinedName newName = new DefinedName(definedName.Workbook, nameId, definedName.Name, definedName.GetAddressString(false), definedName.Scope);
                AddDefinedNameToCollection(newName);
                return this[nameId];
            }
        }

        public void Delete(int index)
        {
            if (Contains(index))
                NamesDictionary.Remove(index);
        }

        // Implement IEnumerable
        public IEnumerator<DefinedName> GetEnumerator()
        {
            return new GenericEnumerator<DefinedName>(NamesDictionary);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }


        /***********************************
         * INTERNAL METHODS
         ************************************/

        internal DefinedName this[int index]
        {
            get
            {
                if (Contains(index))
                    return NamesDictionary[index];
                else
                    return null;
            }
        }

        internal void DeleteWorksheetNames(Worksheet worksheet)
        {
            foreach (DefinedName definedName in this.Where(n => n.Worksheet.Index == worksheet.Index))
            {
                definedName.Delete();
            }
        }

        // Print names should demand strongly typed addresses rather than accepting a string. At the very least there should be some validation.
        internal DefinedName SetWorksheetPrintTitles(Worksheet worksheet, string localAddress)
        {
            DefinedName definedName =
                    new DefinedName(Workbook, -1, PrintTitlesName, worksheet.GetFullAddress(localAddress), DefinedNameScope.Worksheet);
            return Workbook.DefinedNames.Insert(definedName);
        }
        internal DefinedName SetWorksheetPrintArea(Worksheet worksheet, string localAddress)
        {
            DefinedName definedName =
                    new DefinedName(Workbook, -1, PrintAreaName, worksheet.GetFullAddress(localAddress), DefinedNameScope.Worksheet);
            return Workbook.DefinedNames.Insert(definedName);
        }

        internal string GetWorksheetPrintTitlesDefinedName(Worksheet worksheet)
        {
            DefinedName definedName = this.FirstOrDefault(n => n.Worksheet.Index == worksheet.Index && n.Name == PrintTitlesName);
            if (definedName != null)
                return definedName.GetAddressString(true);
            return "";
        }
        internal string GetWorksheetPrintAreaDefinedName(Worksheet worksheet)
        {
            DefinedName definedName = this.FirstOrDefault(n => n.Worksheet.Index == worksheet.Index && n.Name == PrintAreaName);
            if (definedName != null)
                return definedName.GetAddressString(true);
            return "";
        }

        /***********************************
         * PRIVATE METHODS
         ************************************/

        private int GenerateNewNameId()
        {
            return OfficeUtilities.GetFirstUnusedKeyFromCollection<DefinedName>(NamesDictionary);
        }

        private void AddDefinedNameToCollection(DefinedName definedName)
        {
            if (definedName != null)
            {
                if (!Contains(definedName.NameId))
                    NamesDictionary.Add(definedName.NameId, definedName);
            }
        }


        /***********************************
         * DAL METHODS
         ************************************/

        // Read

        internal static DefinedNameCollection ReadDefinedNamesFromReader(CustomOpenXmlReader reader, Workbook workbook)
        {
            DefinedNameCollection definedNames = new DefinedNameCollection(workbook);

            while (reader.ReadToEndElement<OpenXmlSpreadsheet.DefinedNames>())
            {
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.DefinedName>())
                {
                    DefinedName definedName = DefinedName.ReadDefinedNameFromReader(reader, workbook, definedNames.Count + DefinedName.DefaultDefinedNameId);
                    definedNames.AddDefinedNameToCollection(definedName);
                }
            }

            return definedNames;
        }

        // Write
        
        internal static void AddDefinedNamesToWorkbookElement(OpenXmlSpreadsheet.Workbook workbookElement, IEnumerable<DefinedName> definedNames)
        {
            if (workbookElement.DefinedNames != null)
                workbookElement.DefinedNames.Remove();

            if (definedNames.Count() > 0)
            {
                workbookElement.DefinedNames = new OpenXmlSpreadsheet.DefinedNames();
                foreach (DefinedName definedName in definedNames)
                {
                    DefinedName.AddDefinedNameToDefinedNamesElement(workbookElement.DefinedNames, definedName);
                }
            }
        }
        
    }
}
