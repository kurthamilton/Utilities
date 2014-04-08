using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{
    public class Worksheet : BaseExcel, IEquatable<Worksheet>
    {
        public Workbook Workbook { get; private set; }

        private int _index;
        public int Index { get { return _index; } internal set { _index = value; } }
        private string _name;
        public string Name { get { return _name; } set { _name = Workbook.Worksheets.GetUniqueWorksheetName(value); } }        
        private string _relationshipId;
        internal string RelationshipId { get { return _relationshipId; } set { if (value != "") _relationshipId = value; } }
        private WorksheetVisibility _visibility = WorksheetVisibility.Visible;
        public WorksheetVisibility Visibility { get { return _visibility; } set { if (value != WorksheetVisibility.Visible || Workbook.Worksheets.VisibleWorksheetCount > 1) _visibility = value; } }

        /// <summary>
        /// Temporary property. Set using the range address. Not implemented if print setup is not pre-defined.
        /// </summary>
        public string PrintTitles 
        {
            get { return Workbook.DefinedNames.GetWorksheetPrintTitlesDefinedName(this); } 
            set  { Workbook.DefinedNames.SetWorksheetPrintTitles(this, value); } 
        }
        public string PrintArea
        {
            get { return Workbook.DefinedNames.GetWorksheetPrintAreaDefinedName(this); }
            set { Workbook.DefinedNames.SetWorksheetPrintArea(this, value); }
        }

        public string UsedRange
        {
            // hard coded A1 is a bit of a cheat, and not really reflecting the true used area, but it returns a more practical area. Need to compare to Excel's implementation
            get { return string.Format("A1:{0}{1}", Columns.Last().Name, Rows.Last().Index); }
        }

        /// <summary>
        /// Gets or sets the flag to determine whether or not the Worksheet is the active Worksheet in the Workbook. Only one Worksheet can be selected.
        /// </summary>
        public bool Selected 
        { 
            get 
            {
                if (_worksheetData != null)
                    return WorksheetData.SheetViews.Selected;
                else if (Workbook.ActiveTab > 0)
                    return (Index == Workbook.ActiveTab + 1);
                else
                    return (Index == 1);
            } 
            set 
            {
                if (value == true)
                {
                    foreach (Worksheet worksheet in Workbook.Worksheets)
                    {
                        worksheet.Selected = false;
                    }
                }
                WorksheetData.SheetViews.Selected = value; 
                Workbook.ActiveTab = Index - 1; 
            } 
        }

        public int FrozenRow { get { return WorksheetData.SheetViews.FrozenRow; } set { WorksheetData.SheetViews.FrozenRow = value; } }
        public int FrozenColumn { get { return WorksheetData.SheetViews.FrozenColumn; } set { WorksheetData.SheetViews.FrozenColumn = value; } }

        internal string RowSpans { get { return string.Format("1:{0}", Columns.Count); } }

        private WorksheetCellCollection _cells;
        public WorksheetCellCollection Cells { get { if (_cells == null) _cells = new WorksheetCellCollection(this); return _cells; } }        

        // worksheet data 

        private WorksheetData _worksheetData;
        internal WorksheetData WorksheetData 
        {
            get { if (_worksheetData == null) { _worksheetData = new WorksheetData(this); _worksheetData.LoadWorksheetData(); } return _worksheetData; } 
            private set { _worksheetData = value; } 
        }

        public WorksheetFormat Format { get { return WorksheetData.Format; } }
        public ColumnCollection Columns { get { return WorksheetData.Columns; } }
        public RowCollection Rows { get { return WorksheetData.Rows; } }
        public PrintOptions PrintOptions { get { return WorksheetData.PrintOptions; } }
        public PageMargins PageMargins { get { return WorksheetData.PageMargins; } }
        public PageSetup PageSetup { get { return WorksheetData.PageSetup; } }
        public HeaderFooter HeaderFooter { get { return WorksheetData.HeaderFooter; } }
        public Drawing Drawing { get { return WorksheetData.Drawing; } }

        /***********************************
         * CONSTRUCTORS
         ************************************/

        private Worksheet(Workbook workbook)
        {
            Workbook = workbook;
        }

        public Worksheet(Workbook workbook, int index)
            : this(workbook, index, "", WorksheetVisibility.Visible, "")
        {
        }

        public Worksheet(Workbook workbook, int index, string name)
            : this(workbook, index, name, WorksheetVisibility.Visible, "")
        {
        }
        
        internal Worksheet(Workbook workbook, int index, string name, WorksheetVisibility visibility, string relationshipId)
            : this (workbook)
        {
            index = workbook.Worksheets.GetValidWorksheetInsertIndex(index);
            if (name == "")
                name = string.Format("Sheet{0}", index);
            Name = name;
            Index = index;
            _visibility = visibility;
            
            if (relationshipId == "")
                relationshipId = Worksheet.CreateWorksheetPartByWorksheet(this);            
            _relationshipId = relationshipId;            
        }        


        /***********************************
         * PUBLIC PROPERTIES
         ************************************/

        private double _maximumColumnWidth = 0;

        public double MaximumColumnWidth
        {
            get 
            {
                return _maximumColumnWidth; 
            }
            set 
            {
                if (value >= 0)
                {
                    _maximumColumnWidth = value;
                }
            }
        }

        /// <summary>
        /// Get an escaped version of the name, primarily for use as a DefinedName prefix.
        /// </summary>
        internal string EscapedName
        {
            get
            {
                return string.Format("'{0}'", Name.Replace("'", "''"));
            }
        }

        /***********************************
         * PUBLIC METHODS
         ************************************/

        /// <summary>
        /// Adds a data column to the worksheet at the given location.
        /// </summary>
        public void AddDataColumn(DataColumn dataColumn, int rowIndex, int columnIndex, bool addHeader = true, string nullValue = "")
        {
            if (addHeader)
            {
                Cells[rowIndex, columnIndex].Value = dataColumn.ColumnName;
                rowIndex++;
            }

            for (int r = 0; r < dataColumn.Table.Rows.Count; r++)
            {
                if (dataColumn.Table.Rows[r][dataColumn].GetType() == typeof(DateTime) && dataColumn.Table.Rows[r][dataColumn] != null)
                {
                    Cells[rowIndex, columnIndex].Value = string.Format("{0:dd/MM/yyyy HH:mm}", dataColumn.Table.Rows[r][dataColumn]);
                }
                else
                {
                    Cells[rowIndex, columnIndex].Value = dataColumn.Table.Rows[r][dataColumn];
                }

                rowIndex++;
            }
        }

        /// <summary>
        /// Creates a copy of the Worksheet with the given index and name.
        /// </summary>
        public Worksheet CloneTo(Worksheet newWorksheet)
        {
            newWorksheet.WorksheetData = new WorksheetData(newWorksheet);            
            WorksheetData.CloneToWorksheet(newWorksheet);
            return newWorksheet;
        }

        /// <summary>
        /// Copies and inserts the Worksheet to the given Worksheet index.
        /// </summary>
        /// <param name="insertIndex">The insert location for the new Worksheet. Inserts at the beginning or end if index is out of range.</param>
        /// <param name="name">The name of the new Worksheet. Will be made unique if the name already exists in the Workbook.</param>
        public Worksheet CopyTo(int insertIndex, string name)
        {
            Worksheet newWorksheet = new Worksheet(Workbook, insertIndex, name, Visibility, "");
            Workbook.Worksheets.Insert(newWorksheet);
            this.CloneTo(newWorksheet);
            return newWorksheet;
        }

        /// <summary>
        /// Copies and inserts the Worksheet to the given Worksheet index.
        /// </summary>
        /// <param name="insertIndex">The insert location for the new Worksheet. Inserts at the beginning or end if index is out of range.</param>
        public Worksheet CopyTo(int insertIndex)
        {
            return CopyTo(insertIndex, "");
        }

        /// <summary>
        /// Exports the active range of the current Worksheet into a DataTable. Default column headers are the Excel column names. 
        /// Otherwise use the first non-blank row as the header row.
        /// </summary>
        /// <param name="includeRowColumn">Include row numbers as first column?</param>
        public DataTable ToDataTable(bool useCellValueHeaders = false, bool includeRowColumn = true)
        {
            if (useCellValueHeaders)
                return ToDataTableWithCellValueHeaders(includeRowColumn);
            else
                return ToDataTableWithColumnNameHeaders(includeRowColumn);
        }

        /// <summary>
        /// Delete the Worksheet. 
        /// <para></para>Exceptions:<para></para>Exception(Cannot delete last worksheet).
        /// </summary>
        public void Delete()
        {
            Workbook.Worksheets.Delete(Index);
        }

        public string GetLocalAddress(Cell cell1, Cell cell2, bool fixedReference = false)
        {
            return BaseRange.GetRangeAddress(cell1, cell2, fixedReference);
        }
        public string GetLocalAddress(Row row1, Row row2, bool fixedReference = false)
        {
            return BaseRange.GetRangeAddress(row1, row2, fixedReference);
        }
        public string GetLocalAddress(Column column1, Column column2, bool fixedReference = false)
        {
            return BaseRange.GetRangeAddress(column1, column2, fixedReference);
        }

        // implement IEquatable
        public bool Equals(Worksheet other)
        {
            return (other.Workbook.Equals(Workbook) &&
                other.Index == Index);
        }

        /***********************************
         * INTERNAL METHODS
         ************************************/

        /// <summary>
        /// Adds a data table to the worksheet with the given template cells. This called by the PlainReport class. There is a simplified public method.
        /// </summary>
        //internal void AddDataTable(DataTable dataTable, Cell dateTemplateCell, Cell timeTemplateCell,
        //    Cell dateTimeTemplateCell, Cell highlightTemplateCell, string nullValue, bool freezeHeaderRow)
        //{
        //    int rowIndex = PlainReport.GetTemplateRows(this).Max(r => r.Index) + 1;
        //    int columnIndex = 1;

        //    AddDataTableHeader(dataTable, rowIndex, columnIndex);

        //    if (freezeHeaderRow)
        //        FrozenRow = 2;

        //    AddDataTableContent(dataTable, dateTemplateCell, timeTemplateCell, dateTimeTemplateCell, highlightTemplateCell, rowIndex + 1, columnIndex, nullValue);

        //    PlainReport.DeleteTemplateRows(this);
        //}

        internal static string GetLegalWorksheetName(string name)
        {            
            string illegalCharacters = @"/\:[]?*";                        
            foreach (char c in illegalCharacters)
            {
                name = name.Replace(c, '_');
            }

            // name cannot start or end with a apostrophe
            if (name.StartsWith("'"))
                name = name.Substring(1);
            if (name.EndsWith("'"))
                name = name.Substring(0, name.Length - 1);

            // "history" is a reserved worksheet name
            if (string.Compare(name, "history", StringComparison.InvariantCultureIgnoreCase) == 0)
                name = name + " ";

            // names can only be up to 31 characters in length
            if (name.Length > 31)
                name = name.Substring(0, 31);
            
            return name;
        }        

        internal void Save()
        {
            if (_worksheetData != null)
                WorksheetData.Save();
        }

        internal string GetFullAddress(string localAddressString)
        {
            string escapedName = EscapedName;
            
            string[] localAddresses = localAddressString.Split(',');
            
            List<string> addresses = new List<string>();
            foreach (string localAddress in localAddresses)
            {
                addresses.Add(string.Format("{0}!{1}", escapedName, localAddress));
            }

            return string.Join(",", addresses.ToArray());
        }

        /***********************************
         * PRIVATE METHODS
         ************************************/

        private DataTable ToDataTableWithColumnNameHeaders(bool includeRowColumn)
        {
            DataTable dataTable = CreateDataTableForExport(includeRowColumn);

            foreach (Column column in Columns)
            {
                dataTable.Columns.Add(new DataColumn(column.Name, typeof(string)));
            }

            foreach (Row row in Rows)
            {
                DataRow dataRow = dataTable.NewRow();

                if (includeRowColumn)
                    dataRow[0] = row.Index;

                row.Cells.Action(c => dataRow[c.Column.Name] = c.Text);

                dataTable.Rows.Add(dataRow);
            }

            return dataTable;
        }

        private DataTable ToDataTableWithCellValueHeaders(bool includeRowColumn)
        {
            DataTable dataTable = CreateDataTableForExport(includeRowColumn);

            Row headerRow = Rows.GetFirstNonBlankRow();

            if (headerRow != null)
            {
                List<Cell> headerCells = GetHeaderCellsForDataTable(headerRow);

                if (CheckHeaderRowIsValidForColumnNames(headerCells))
                {
                    // add columns
                    foreach (Cell cell in headerCells)
                    {
                        string columnName = cell.Text.Trim();

                        dataTable.Columns.Add(new DataColumn(columnName, typeof(string)));
                    }

                    // get last row
                    int lastRowIndex = headerRow.Index;
                    foreach (Row row in Rows)
                    {
                        if (row.Index > headerRow.Index)
                        {
                            if (row.Cells.GetFirstNonBlankCell() == null)
                                break;

                            lastRowIndex = row.Index;
                        }
                    }

                    // add rows
                    foreach (Row row in Rows)
                    {
                        if (row.Index > headerRow.Index)
                        {
                            if (row.Index > lastRowIndex)
                                break;

                            DataRow dataRow = dataTable.NewRow();

                            int columnIndex = 0;

                            if (includeRowColumn)
                            {
                                dataRow[columnIndex] = row.Index;
                                columnIndex++;
                            }

                            foreach (Cell cell in headerCells)
                            {
                                dataRow[columnIndex] = row.Cells[cell.Column.Index].Text;
                                columnIndex++;
                            }

                            dataTable.Rows.Add(dataRow);
                        }
                    }
                }
            }

            return dataTable;
        }

        private DataTable CreateDataTableForExport(bool includeRowColumn)
        {
            DataTable dataTable = new DataTable("WorksheetData");
            if (includeRowColumn)
                dataTable.Columns.Add(new DataColumn("Row", typeof(int)));
            return dataTable;
        }

        private List<Cell> GetHeaderCellsForDataTable(Row headerRow)
        {
            List<Cell> cells = new List<Cell>();

            int columnIndex = 1;
            while (headerRow.Cells[columnIndex].Text != "")
            {
                cells.Add(headerRow.Cells[columnIndex]);
                columnIndex++;
            }

            return cells;
        }

        protected bool CheckHeaderRowIsValidForColumnNames(List<Cell> headerCells)
        {
            List<string> columnNames = new List<string>();

            foreach (Cell cell in headerCells)
            {
                string columnName = cell.Text.Trim().ToLower();

                if (!columnNames.Contains(columnName))
                    columnNames.Add(columnName);
                else
                    return false;
            }

            return (columnNames.Count > 0);
        }


        private void AddDataTableHeader(DataTable dataTable, int rowIndex, int columnIndex)
        {
            if (dataTable != null)
            {
                Row headerRow = Rows.Insert(rowIndex);

                foreach (DataColumn column in dataTable.Columns)
                {
                    Cell headerCell = headerRow.Cells[columnIndex];
                    headerCell.Value = column.ColumnName;
                    headerCell.Font.Bold = true;
                    columnIndex++;
                }
            }
        }

        private void AddDataTableContent(DataTable dataTable, Cell dateTemplateCell, Cell timeTemplateCell,
            Cell dateTimeTemplateCell, Cell highlightTemplateCell, int rowIndex, int startColumnIndex, string nullValue)
        {
            if (dataTable != null)
            {
                Row dataTemplateRow = dateTemplateCell.Row;

                foreach (DataRow row in dataTable.Rows)
                {
                    Row dataRow = Rows[rowIndex + 1];

                    int columnIndex = startColumnIndex;
                    foreach (DataColumn column in dataTable.Columns)
                    {
                        Cell templateCell = null;
                        object value = row[column];

                        if (column.DataType == typeof(DateTime))
                        {
                            if (value == DBNull.Value || value == null)
                                templateCell = dateTimeTemplateCell;
                            else
                            {
                                DateTime dateTimeValue = (DateTime)value;

                                if (dateTimeValue.TimeOfDay == TimeSpan.Zero)
                                    templateCell = dateTemplateCell;
                                else if (dateTimeValue.Date <= new DateTime(1900, 1, 1))
                                    templateCell = timeTemplateCell;
                                else
                                    templateCell = dateTimeTemplateCell;
                            }
                        }
                        else if (row[column].ToString().Contains("**HIGHLIGHT**"))
                            templateCell = highlightTemplateCell;

                        Cell newCell;
                        if (templateCell != null)
                            newCell = templateCell.CopyTo(rowIndex, columnIndex);
                        else
                            newCell = Cells[rowIndex, columnIndex];

                        if (value == DBNull.Value || value == null)
                            newCell.Value = nullValue;
                        else
                            newCell.Value = value;

                        if (templateCell == highlightTemplateCell)
                            newCell.Value = newCell.Value.ToString().Replace("**HIGHLIGHT**", "");

                        columnIndex++;
                    }

                    rowIndex++;
                }
            }
        }

        /***********************************
         * DAL METHODS
         ************************************/


        // Read
        internal static Worksheet ReadWorksheetFromReader(CustomOpenXmlReader reader, Workbook workbook, int index)
        {
            Worksheet worksheet = new Worksheet(workbook);
            worksheet._index = index;

            // This method sets the private variables to bypass the Workbook.Worksheets (which don't exist yet) logic checks
            foreach (CustomOpenXmlAttribute attribute in reader.Attributes)
            {
                switch (attribute.LocalName)
                {
                    case "name":
                        worksheet._name = attribute.Value;
                        break;
                    case "id":
                        worksheet._relationshipId = attribute.Value;
                        break;
                    case "state":
                        string state = attribute.Value;
                        if (state != "")
                            worksheet._visibility = Helpers.GetEnumValueFromDescription<WorksheetVisibility>(state);
                        break;
                    case "sheetId":
                        // it turns out sheetId can be any value.
                        //worksheet._index = attribute.GetIntValue();
                        break;
                }
            }

            return worksheet;
        }

        private static List<OpenXmlSpreadsheet.Sheet> GetSheetElementsFromDocument(OpenXmlPackaging.SpreadsheetDocument document)
        {
            List<OpenXmlSpreadsheet.Sheet> sheetElements = new List<OpenXmlSpreadsheet.Sheet>();
            foreach (OpenXmlSpreadsheet.Sheet sheet in document.WorkbookPart.Workbook.Sheets.ToList())
            {
                sheetElements.Add(sheet);
            }
            return sheetElements;
        }
        private static OpenXmlSpreadsheet.Sheet GetSheetElementFromWorksheet(Worksheet worksheet)
        {
            List<OpenXmlSpreadsheet.Sheet> sheetElements = GetSheetElementsFromDocument(worksheet.Workbook.Document);
            return sheetElements.Find(s => s.Id == worksheet.RelationshipId);
        }

        internal static OpenXmlSpreadsheet.Worksheet GetWorksheetElementFromWorksheet(Worksheet worksheet)
        {
            OpenXmlSpreadsheet.Sheet sheetElement = GetSheetElementFromWorksheet(worksheet);

            if (sheetElement != null)
            {
                OpenXmlPackaging.WorksheetPart worksheetPart = Worksheet.GetWorksheetPartByWorksheet(worksheet);
                OpenXmlSpreadsheet.Worksheet worksheetElement = worksheetPart.Worksheet;
                return worksheetElement;
            }
            return null;
        }

        internal static OpenXmlPackaging.WorksheetPart GetWorksheetPartByWorksheet(Worksheet worksheet)
        {
            OpenXmlPackaging.SpreadsheetDocument document = worksheet.Workbook.Document;
            string relationshipId = worksheet.RelationshipId;
            return (OpenXmlPackaging.WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
        }

        // Write
        
        internal static void AddWorksheetToSheetsElement(OpenXmlSpreadsheet.Sheets sheetsElement, Worksheet worksheet)
        {
            OpenXmlSpreadsheet.Sheet sheetElement = new OpenXmlSpreadsheet.Sheet();
            sheetElement.Name = worksheet.Name;
            sheetElement.Id = worksheet.RelationshipId;
            sheetElement.State = new OpenXml.EnumValue<OpenXmlSpreadsheet.SheetStateValues>((OpenXmlSpreadsheet.SheetStateValues)((int)worksheet.Visibility));
            sheetElement.SheetId = new OpenXml.UInt32Value((UInt32)worksheet.Index);
            sheetsElement.Append(sheetElement);
        }

        internal static void AddWorksheetPartByWorksheet(Worksheet worksheet)
        {
            OpenXmlPackaging.SpreadsheetDocument spreadSheet = worksheet.Workbook.Document;

            // Add a blank WorksheetPart.
            OpenXmlPackaging.WorksheetPart newWorksheetPart = spreadSheet.WorkbookPart.AddNewPart<OpenXmlPackaging.WorksheetPart>();
            newWorksheetPart.Worksheet = new OpenXmlSpreadsheet.Worksheet(new OpenXmlSpreadsheet.SheetData());

            OpenXmlSpreadsheet.Sheets sheets = spreadSheet.WorkbookPart.Workbook.GetFirstChild<OpenXmlSpreadsheet.Sheets>();
            string relationshipId = spreadSheet.WorkbookPart.GetIdOfPart(newWorksheetPart);
            
            // Get a unique ID for the new worksheet.
            uint sheetId = 1;
            if (sheets.Elements<OpenXmlSpreadsheet.Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<OpenXmlSpreadsheet.Sheet>().Max(s => s.SheetId.Value) + 1;
            }

            // Give the new worksheet a name.
            string sheetName = "Sheet" + sheetId;

            // Append the new worksheet and associate it with the workbook.
            OpenXmlSpreadsheet.Sheet sheet = new OpenXmlSpreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);

            worksheet.RelationshipId = relationshipId;
        }

        internal static void DeleteWorksheetPartByWorksheet(Worksheet worksheet)
        {
            var worksheetElement = GetWorksheetElementFromWorksheet(worksheet);
            if (worksheetElement != null)
            {
                var worksheetPart = worksheetElement.WorksheetPart;

                if (worksheetPart != null)
                {
                    worksheet.Workbook.Document.WorkbookPart.DeletePart(worksheetPart);

                    OpenXmlSpreadsheet.Sheet sheetElement = GetSheetElementFromWorksheet(worksheet);
                    if (sheetElement != null)
                    {
                        sheetElement.Remove();
                        worksheet.Workbook.Document.WorkbookPart.Workbook.Save();
                    }
                }
            }
        }

        private static string CreateWorksheetPartByWorksheet(Worksheet worksheet)
        {
            OpenXmlPackaging.WorkbookPart workbookPart = worksheet.Workbook.Document.WorkbookPart;
            OpenXmlSpreadsheet.Worksheet existingWorksheetElement = workbookPart.WorksheetParts.First().Worksheet;
            OpenXmlPackaging.WorksheetPart worksheetPart = OpenXmlUtilities.CreatePart<OpenXmlPackaging.WorkbookPart, OpenXmlPackaging.WorksheetPart>(workbookPart);

            OpenXmlSpreadsheet.Worksheet worksheetElement = new OpenXmlSpreadsheet.Worksheet();
            // Copy namespace declarations from existing instance. There is probably a better way to do this.
            foreach (var namespaceDeclaration in existingWorksheetElement.NamespaceDeclarations)
            {
                worksheetElement.AddNamespaceDeclaration(namespaceDeclaration.Key, namespaceDeclaration.Value);
            }

            worksheetPart.Worksheet = worksheetElement;
            
            return workbookPart.GetIdOfPart(worksheetPart);
        }
    }
}
