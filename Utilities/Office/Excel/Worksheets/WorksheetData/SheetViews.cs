using System;
using System.Collections.Generic;

using System.Xml;
using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{
    public class SheetViews : BaseExcel
    {
        public Worksheet Worksheet { get; private set; }
        public bool Selected { get; set; }
        public int FrozenRow { get; set; }
        public int FrozenColumn { get; set; }

        /***********************************
         * CONSTRUCTORS
         ************************************/        


        public SheetViews(Worksheet worksheet)
        {
            Worksheet = worksheet;
        }

        public SheetViews(Worksheet worksheet, bool selected, int frozenRow, int frozenColumn)
            : this (worksheet)
        {
            Selected = selected;
            FrozenRow = frozenRow;
            FrozenColumn = frozenColumn;
        }

        /***********************************
         * INTERNAL METHODS
         ************************************/

        internal SheetViews Clone(Worksheet worksheet)
        {
            SheetViews newSheetViews = new SheetViews(worksheet, false, FrozenRow, FrozenColumn);
            return newSheetViews;
        }

        internal bool HasValue()
        {
            return (Selected || FrozenRow > 0 || FrozenColumn > 0);
        }

        /***********************************
         * DAL METHODS
         ************************************/

        // Read

        internal static SheetViews ReadSheetViewsFromReader(CustomOpenXmlReader reader, Worksheet worksheet)
        {
            SheetViews sheetViews = new SheetViews(worksheet);

            while (reader.ReadToEndElement<OpenXmlSpreadsheet.SheetViews>())
            {
                // only read the first sheetView element.
                if (reader.IsEndElementOfType<OpenXmlSpreadsheet.SheetView>())
                    break;
                
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.SheetView>())
                {
                    foreach (CustomOpenXmlAttribute attribute in reader.Attributes)
                    {
                        switch (attribute.LocalName)
                        {
                            case "tabSelected":
                                sheetViews.Selected = attribute.GetBoolValue();
                                break;
                        }
                    }
                }
                else if (reader.IsStartElementOfType<OpenXmlSpreadsheet.Pane>())
                {
                    int ySplit = 0;
                    int xSplit = 0;
                    bool isFrozen = false;

                    foreach (CustomOpenXmlAttribute attribute in reader.Attributes)
                    {
                        switch (attribute.LocalName)
                        {
                            case "ySplit":
                                ySplit = attribute.GetIntValue();
                                break;
                            case "xSplit":
                                xSplit = attribute.GetIntValue();
                                break;
                            case "state":
                                isFrozen = (attribute.Value == "frozen");
                                break;
                        }
                    }

                    if (isFrozen)
                    {
                        if (ySplit > 0) sheetViews.FrozenRow = ySplit;
                        if (xSplit > 0) sheetViews.FrozenRow = xSplit;
                    }
                }

                
            }

            return sheetViews;
        }

        // Write

        internal static void WriteSheetViewsToWriter(CustomOpenXmlWriter<OpenXmlPackaging.WorksheetPart> writer, SheetViews sheetViews)
        {
            if (sheetViews.HasValue())
            {
                writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.SheetViews());
                writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.SheetView());

                if (sheetViews.Selected)
                    writer.WriteAttribute("tabSelected", sheetViews.Selected);
                // the implementation of worksheetViewId gets quite complicated. The use isn't supported here, but the attribute is necessary to compile.
                writer.WriteAttribute("workbookViewId", 0);

                if (sheetViews.FrozenRow > 0 || sheetViews.FrozenColumn > 0)
                {
                    writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.Pane());
                    writer.WriteAttribute("state", "frozen");
                    if (sheetViews.FrozenRow > 0) writer.WriteAttribute("ySplit", sheetViews.FrozenRow);
                    if (sheetViews.FrozenColumn > 0) writer.WriteAttribute("xSplit", sheetViews.FrozenColumn);
                    writer.WriteAttribute("topLeftCell", sheetViews.Worksheet.Cells[sheetViews.FrozenRow + 1, sheetViews.FrozenColumn + 1].Address);
                    writer.WriteEndElement();   // Pane
                }

                writer.WriteEndElement();   // SheetView
                writer.WriteEndElement();   // SheetViews
            }
        }
    }
}
