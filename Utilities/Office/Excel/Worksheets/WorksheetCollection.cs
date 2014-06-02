using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{
    public class WorksheetCollection : BaseExcel, IEnumerable<Worksheet>
    {
        public const int MinValue = 1;

        private Workbook Workbook { get; set; }

        private SortedDictionary<int, Worksheet> WorksheetDictionary { get; set; }
        
        public List<Worksheet> Worksheets { get { return WorksheetDictionary.Values.ToList(); } }
        public int Count { get { return WorksheetDictionary.Count; } }
        
        public List<Worksheet> VisibleWorksheets { get { return GetVisibleWorksheets(); } }
        public int VisibleWorksheetCount { get { return VisibleWorksheets.Count;  } }
        
        /// <summary>
        /// Gets the currently selected Worksheet in the Workbook's Worksheet collection.
        /// </summary>
        public Worksheet ActiveWorksheet { get { if (GetActiveWorksheets().Count == 0) WorksheetDictionary[1].Selected = true; return GetActiveWorksheets()[0]; } }
        

        /***********************************
         * CONSTRUCTORS
         ************************************/

        internal WorksheetCollection(Workbook workbook)
        {
            Workbook = workbook;
            WorksheetDictionary = new SortedDictionary<int, Worksheet>();
        }
        

        /***********************************
         * PROTECTED METHODS
         ************************************/

        protected void AddWorksheetToCollection(Worksheet worksheet)
        {
            WorksheetDictionary.Add(worksheet.Index, worksheet);
        }

        protected void RemoveWorksheetFromCollection(Worksheet worksheet)
        {
            if (Contains(worksheet.Index))
                WorksheetDictionary.Remove(worksheet.Index);
        }


        /***********************************
         * PUBLIC METHODS
         ************************************/
        
               
        public Worksheet this[int index]
        {
            get
            {
                if (Contains(index))
                    return WorksheetDictionary[index];
                else
                    throw new ArgumentOutOfRangeException();
            }

            set
            {
                if (index > 0 && index <= Worksheets.Count + 1)
                {
                    if (Contains(index))
                        WorksheetDictionary[index] = value;
                    else
                        AddWorksheetToCollection(value);
                }
                else
                    throw new ArgumentOutOfRangeException();
            }
        }

        public Worksheet this[string name]
        {
            get
            {
                if (Contains(name))
                    return WorksheetDictionary.Values.First(ws => string.Compare(name, ws.Name, true) == 0);
                else
                    throw new ArgumentOutOfRangeException();
            }
        }

        /// <summary>
        /// Determines whether the given Worksheet index exists within the Workbook Worksheets.
        /// </summary>
        public bool Contains(int index)
        {
            return WorksheetDictionary.ContainsKey(index);
        }

        /// <summary>
        /// Determines whether the named Worksheet exists within the Workbook Worksheets.
        /// </summary>
        public bool Contains(string name)
        {
            return (WorksheetDictionary.Values.FirstOrDefault(ws => string.Compare(name, ws.Name, true) == 0) != null);
        }

        /// <summary>
        /// Add a new blank Worksheet after the last existing Worksheet.
        /// </summary>
        public Worksheet Add(string name)
        {
            Worksheet newWorksheet = new Worksheet(Workbook, Workbook.Worksheets.Count + 1, name);
            AddWorksheetToCollection(newWorksheet);
            
            // this should not be mixed in with the business logic like this.
            Worksheet.AddWorksheetPartByWorksheet(newWorksheet);

            return newWorksheet;
        }        

        /// <summary>
        /// Insert a new blank Worksheet with the given name at the given index.
        /// </summary>
        public Worksheet Insert(int insertIndex, string name)
        {
            Worksheet newWorksheet = new Worksheet(Workbook, insertIndex, name);
            return Insert(newWorksheet);
        }

        /// <summary>
        /// Insert a new blank Worksheet at the given index.
        /// </summary>
        public Worksheet Insert(int insertIndex)
        {
            return Insert(insertIndex, string.Empty);
        }


        /// <summary>
        /// Delete the Worksheet with the given index.
        /// <para></para>Exceptions:<para></para>Exception(Cannot delete last worksheet).
        /// </summary>
        public void Delete(int index)
        {
            if (Count == 1)
                throw new Exception("Cannot delete last worksheet");

            if (Contains(index))
            {
                Worksheet deletedWorksheet = this[index];

                Workbook.DefinedNames.DeleteWorksheetNames(deletedWorksheet);

                // this should not be mixed in with the business logic like this.
                Worksheet.DeleteWorksheetPartByWorksheet(deletedWorksheet);

                IEnumerable<Worksheet> affectedWorksheets = Worksheets.Where(w => w.Index >= index);

                foreach (Worksheet worksheet in affectedWorksheets)
                {
                    RemoveWorksheetFromCollection(worksheet);

                    if (worksheet.Index != index)
                    {
                        worksheet.Index--;
                        AddWorksheetToCollection(worksheet);
                    }
                }
            }
        }

        // Implement IEnumerable
        public IEnumerator<Worksheet> GetEnumerator()
        {
            return new GenericEnumerator<Worksheet>(WorksheetDictionary);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        /***********************************
         * INTERNAL METHODS
         ************************************/

        internal int GetValidWorksheetInsertIndex(int attemptedInsertIndex)
        {
            if (Workbook.HasLoaded)
            {
                if (attemptedInsertIndex < 1)
                    attemptedInsertIndex = 1;
                else if (attemptedInsertIndex > Worksheets.Count + 1)
                    attemptedInsertIndex = Worksheets.Count + 1;
            }

            return attemptedInsertIndex;
        }

        internal string GetUniqueWorksheetName(string name)
        {
            name = Worksheet.GetLegalWorksheetName(name);

            int version = 1;
            while (Worksheets.Find(w => string.Compare(name, w.Name, true) == 0) != null)
            {
                string versionSuffix = string.Empty;
                if (version > 1)
                    versionSuffix = string.Format(" ({0})", version);

                string versionedTabName = string.Concat(name, versionSuffix);
                if (versionedTabName.Length > 31)
                    versionedTabName = string.Concat(name.Substring(0, 31 - versionSuffix.Length), versionSuffix);

                name = versionedTabName;

                version++;
            }

            return name;
        }

        internal Worksheet Insert(Worksheet newWorksheet)
        {
            int insertIndex = newWorksheet.Index;

            List<Worksheet> affectedWorksheets = WorksheetDictionary.Values.Where(w => w.Index >= insertIndex).ToList();
            // update worksheets in reverse order to avoid creating conflicting worksheet index keys when updating worksheet collection
            affectedWorksheets.Reverse();

            foreach (Worksheet worksheet in affectedWorksheets)
            {
                RemoveWorksheetFromCollection(worksheet);
                worksheet.Index++;
                AddWorksheetToCollection(worksheet);
            }

            AddWorksheetToCollection(newWorksheet);

            return WorksheetDictionary[insertIndex];
        }

        internal void Save()
        {
            foreach (Worksheet worksheet in this)
            {
                worksheet.Save();
            }
        }

        /***********************************
         * PRIVATE METHODS
         ************************************/


        private List<Worksheet> GetVisibleWorksheets()
        {
            return WorksheetDictionary.Values.Where(w => w.Visibility == WorksheetVisibility.Visible).ToList();
        }

        private List<Worksheet> GetActiveWorksheets()
        {            
            return WorksheetDictionary.Values.Where(w => w.Selected).ToList();
        }




        /***********************************
        * DAL METHODS
        ************************************/


        // Read

        internal static WorksheetCollection ReadWorksheetsFromReader(CustomOpenXmlReader reader, Workbook workbook)
        {
            WorksheetCollection worksheets = new WorksheetCollection(workbook);

            while (reader.ReadToEndElement<OpenXmlSpreadsheet.Sheets>())
            {
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.Sheet>())
                {
                    Worksheet worksheet = Worksheet.ReadWorksheetFromReader(reader, workbook, worksheets.Count + 1);
                    worksheets.AddWorksheetToCollection(worksheet);
                }
            }

            return worksheets;
        }

        // Write

        internal static void AddWorksheetsToWorkbookElement(OpenXmlSpreadsheet.Workbook workbookElement, IEnumerable<Worksheet> worksheets)
        {
            workbookElement.Sheets.RemoveAllChildren();

            foreach (Worksheet worksheet in worksheets)
            {
                Worksheet.AddWorksheetToSheetsElement(workbookElement.Sheets, worksheet);
            }

            workbookElement.Save();
        }        
    }
}
