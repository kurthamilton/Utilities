using System;
using System.Linq;

using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{
    internal class WorkbookData : BaseExcel
    {
        public Workbook Workbook { get; private set; }

        private int _activeTab;
        public int ActiveTab { get { return _activeTab; } set { _activeTab = value; } }

        private WorksheetCollection _worksheets;
        public WorksheetCollection Worksheets { get { return _worksheets; } set { _worksheets = value; } }

        private DefinedNameCollection _definedNames;
        internal DefinedNameCollection DefinedNames { get { if (_definedNames == null) _definedNames = new DefinedNameCollection(Workbook); return _definedNames; } private set { _definedNames = value; } }

        /***********************************
         * CONSTRUCTORS
         ************************************/

        public WorkbookData(Workbook workbook)
        {
            Workbook = workbook;            
        }


        /***********************************
         * PUBLIC METHODS
         ************************************/

        public void Save()
        {
            SaveWorkbookData(this);
            Worksheets.Save();            
        }

        /***********************************
         * INTERNAL METHODS
         ************************************/

        internal void LoadWorkbookData()
        {            
            WorkbookData.ReadWorkbookDataFromWorkbookData(this);
        }

        /***********************************
         * DAL METHODS
         ************************************/

        // Read

        private static void ReadWorkbookDataFromWorkbookData(WorkbookData workbookData)
        {
            OpenXmlSpreadsheet.Workbook workbookElement = workbookData.Workbook.Document.WorkbookPart.Workbook;

            using (CustomOpenXmlReader reader = CustomOpenXmlReader.Create(workbookElement))
            {
                while (reader.Read())
                {
                    if (reader.IsStartElementOfType<OpenXmlSpreadsheet.WorkbookView>())
                    {
                        CustomOpenXmlAttribute attribute = reader.Attributes["activeTab"];
                        if (attribute != null)
                            workbookData.ActiveTab = attribute.GetIntValue();
                    }
                    else if (reader.IsStartElementOfType<OpenXmlSpreadsheet.Sheets>())
                        workbookData.Worksheets = WorksheetCollection.ReadWorksheetsFromReader(reader, workbookData.Workbook);
                    else if (reader.IsStartElementOfType<OpenXmlSpreadsheet.DefinedNames>())
                        workbookData.DefinedNames = DefinedNameCollection.ReadDefinedNamesFromReader(reader, workbookData.Workbook);
                }
            }
        }

        // Write

        internal static void SaveWorkbookData(WorkbookData workbookData)
        {
            // Use DOM instead of XmlWriter when writing to workbook part and re-creating the workbook part causes problems
            // This means that the workbook xml file is saved with x: prefixed nodes and the file is loaded into memory, but it's not much of an issue on this small file.
            OpenXmlSpreadsheet.Workbook workbookElement = workbookData.Workbook.Document.WorkbookPart.Workbook;

            ((OpenXmlSpreadsheet.WorkbookView)workbookElement.BookViews.FirstChild).ActiveTab = new OpenXml.UInt32Value((UInt32)workbookData.ActiveTab);
            WorksheetCollection.AddWorksheetsToWorkbookElement(workbookElement, workbookData.Worksheets);
            if (workbookData.DefinedNames != null)
                DefinedNameCollection.AddDefinedNamesToWorkbookElement(workbookElement, workbookData.DefinedNames);

            workbookElement.Save();
        }
    }
}
