using System;
using System.IO;

using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;

namespace Utilities.Office.Excel
{
    /*
     * TO DO
     * Read Merge Cells
     * Read Hyperlinks
     * Develop BaseRange to truly support ranges. At the moment it is just a base class for Cell
     * 
     * OLD NOTES
        // - Create new Excel workbook from scratch
        // - Images - can only replace at the moment
        // - Create new shared strings xml file + workbook rel if no shared strings initially exist. This should be easy to do now using the OpenXml SDK.
        // - Support formula cells (currently just removing calc chain rel).
        // - Combine Cell.MergeDown and Cell.MergeAcross.
     */

    public class Workbook : BaseExcel, IDisposable, IEquatable<Workbook>
    {
        /// <summary>
        /// Gets the full network location and file name of the Workbook.
        /// </summary>
        public string FilePath { get; private set; }

        internal bool ReadOnly { get; private set; }
        internal bool HasLoaded { get; private set; }   // try to deprecate
        
        internal OpenXmlPackaging.SpreadsheetDocument Document { get; private set; }

        /// <summary>
        /// Gets the name of the Workbook.
        /// </summary>
        public string Name { get { return new FileInfo(FilePath).Name; } }
        
        /// <summary>
        /// Gets the full network location of the Workbook.
        /// </summary>
        public string Path { get { return new DirectoryInfo(FilePath).Name; } }        

        /// <summary>
        /// Gets the currently selected Worksheet in the Workbook's Worksheet collection.
        /// </summary>
        public Worksheet ActiveWorksheet { get { return Worksheets.ActiveWorksheet; } }
        internal int ActiveTab { get { return WorkbookData.ActiveTab; } set { WorkbookData.ActiveTab = value; } }

        // Workbook data - stored in workbook.xml
        private WorkbookData _workbookData;
        private WorkbookData WorkbookData 
        {
            get { if (_workbookData == null) { _workbookData = new WorkbookData(this); _workbookData.LoadWorkbookData(); } return _workbookData; } 
            set { _workbookData = value; } 
        }

        public WorksheetCollection Worksheets { get { return WorkbookData.Worksheets; } }
        internal DefinedNameCollection DefinedNames { get { return WorkbookData.DefinedNames; } }

        public CellFormat DefaultFormat { get { return Styles.CellFormats[CellFormat.DefaultStyleIndex]; } }

        // Workbook parts   
        private SharedStringCollection _sharedStrings;
        internal SharedStringCollection SharedStrings { get { if (_sharedStrings == null) _sharedStrings = new SharedStringCollection(this); return _sharedStrings; } }

        private Styles _styles;
        internal Styles Styles { get { if (_styles == null) _styles = new Styles(this); return _styles; } }

        //private DrawingCollection _drawings;
        //internal DrawingCollection Drawings { get { if (_drawings == null) _drawings = new DrawingCollection(this); return _drawings; } private set { _drawings = value; } }

        /***********************************
         * CONSTRUCTORS
         ************************************/

        //public Workbook()
        //    //: this(Workbook.CopyBlankTemplate())
        //{
            
        //}

        public Workbook(string filePath)
            : this(filePath, false)
        {
        }

        public Workbook(string filePath, bool readOnly)
        {
            ReadOnly = readOnly;
            FilePath = filePath;

            Open();
        }

        private void Open()
        {
            if (!(File.Exists(FilePath)))
                throw new FileNotFoundException(string.Format("Error opening Excel Workbook. File not found {0}", FilePath));

            if (!FilePath.ToLower().EndsWith(".xlsx"))
                throw new FileFormatException(string.Format("Error opening Excel Workbook. Invalid file format {0}", FilePath));

            try
            {
                Document = OpenXmlPackaging.SpreadsheetDocument.Open(FilePath, !ReadOnly);
            }
            catch (UriFormatException)
            {
                // If "invalid" uris (references to C:\ for example) are found when opening a SpreadsheetDocument an exception is thrown immediately. 
                // Catch invalid uris, clean up, and re-try to open.
                WorkbookCleaner.CleanWorkbook(FilePath);
                Document = OpenXmlPackaging.SpreadsheetDocument.Open(FilePath, !ReadOnly);
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Save all changes made to the Workbook.
        /// </summary>
        public void Save()
        {
            Workbook.SaveWorkbook(this);
            
            // Check the private variables to avoid Loading the main variable if not yet loaded
            if (_workbookData != null)
                WorkbookData.Save();

            if (_sharedStrings != null)
                SharedStrings.Save();

            if (_styles != null)
                Styles.Save();
        }

        /// <summary>
        /// Save the Workbook with all changes to a new File location.
        /// </summary>
        public void SaveAs(string filePath)
        {
            throw new NotImplementedException();          
        }

        /// <summary>
        /// Close the Workbook. 
        /// </summary>
        /// <param name="saveChanges">Whether to save all changes made to the Workbook before closing</param>
        public void Close(bool saveChanges = true)
        {
            if (saveChanges)
                Save();
            Document.Close();            
        }

        // Implement IDisposable
        /// <summary>
        /// <para>Close the Workbook and saves changes if the Workbook is not read-only.</para>
        /// </summary>
        public void Dispose()
        {
            Close(!ReadOnly);
        }

        // implement IEquatable
        public bool Equals(Workbook other)
        {
            return (string.Compare(other.FilePath, FilePath, StringComparison.InvariantCultureIgnoreCase) == 0);
        }        


        /***********************************
         * PRIVATE METHODS
         ************************************/        

        //private static string CopyBlankTemplate()
        //{
        //    return OfficeIO.TryCopyFile(ExcelIO.GetBlankReportTemplateFilePath(), OfficeIO.GetWorkingFolderPath(), "Blank Template.xlsx");
        //}


        /***********************************
         * DAL METHODS
         ************************************/

        // Read

        // Write

        private static void SaveWorkbook(Workbook workbook)
        {
            // Formulae not supported, so remove calc chain if it exists

            OpenXmlPackaging.WorkbookPart workbookPart = workbook.Document.WorkbookPart;
            OpenXmlPackaging.CalculationChainPart calcChainPart = workbookPart.CalculationChainPart;
            if (calcChainPart != null)
                workbookPart.DeletePart(calcChainPart);
        }
    }
}
