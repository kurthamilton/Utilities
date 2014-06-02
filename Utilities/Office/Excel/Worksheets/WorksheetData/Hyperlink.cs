using System;
using System.Collections.Generic;
using System.Linq;

using System.Xml;
using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{
    internal static class Hyperlink
    {

        /***********************************
         * PRIVATE METHODS
         ************************************/

        private static string GetTargetAddress(Cell targetCell)
        {
            return string.Format("{0}!{1}", targetCell.Worksheet.Name, targetCell.Address);
        }

        private static Cell GetTargetCellByTargetLocation(Worksheet worksheet, string targetLocation)
        {
            string[] targetElements = targetLocation.Split('!');

            // not all Hyperlinks are to worksheets. Only handle worksheet hyperlinks.
            if (targetElements.Length >= 2 && worksheet.Workbook.Worksheets.Contains(targetElements[0]))
            {
                Worksheet targetWorksheet = worksheet.Workbook.Worksheets[targetElements[0]];
                Cell targetCell = targetWorksheet.Cells[targetElements[1]];
                return targetCell;
            }
            return null;
        }


        /***********************************
         * DAL METHODS
         ************************************/

        // Read

        internal static void UpdateCellHyperlinksFromReader(CustomOpenXmlReader reader, Worksheet worksheet)
        {
            while (reader.ReadToEndElement<OpenXmlSpreadsheet.Hyperlinks>())
            {
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.Hyperlink>())
                    UpdateCellHyperlinkFromReader(reader, worksheet);
            }
        }

        private static void UpdateCellHyperlinkFromReader(CustomOpenXmlReader reader, Worksheet worksheet)
        {
            string address = "";
            string display = "";
            string location = "";

            foreach (CustomOpenXmlAttribute attribute in reader.Attributes)
            {
                switch (attribute.LocalName)
                {
                    case "ref":
                        address = attribute.Value;
                        break;
                    case "display":
                        display = attribute.Value;
                        break;
                    case "location":
                        location = attribute.Value;
                        break;
                }
            }

            Cell cell = worksheet.Cells[address];
            Cell targetCell = GetTargetCellByTargetLocation(worksheet, location);

            if (cell != null && targetCell != null)
            {
                cell.HyperlinkToCell = targetCell;
                cell.Value = display;
            }
        }

        // Write

        internal static void WriteHyperlinksToWriter(CustomOpenXmlWriter<OpenXmlPackaging.WorksheetPart> writer, Worksheet worksheet)
        {
            List<Cell> hyperlinkCells = new List<Cell>();
            foreach (Row row in worksheet.Rows)
            {
                hyperlinkCells.AddRange(row.Cells.Where(c => c.HyperlinkToCell != null));
            }

            if (hyperlinkCells.Count > 0)
            {
                writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.Hyperlinks());

                foreach (Cell cell in hyperlinkCells)
                {
                    Hyperlink.WriteHyperlinkToWriter(writer, cell);
                }

                writer.WriteEndElement();   // Hyperlinks
            }
        }

        private static void WriteHyperlinkToWriter(CustomOpenXmlWriter<OpenXmlPackaging.WorksheetPart> writer, Cell cell)
        {
            writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.Hyperlink());

            writer.WriteAttribute("ref", cell.Address);
            writer.WriteAttribute("display", cell.Text);
            writer.WriteAttribute("location", GetTargetAddress(cell.HyperlinkToCell));

            writer.WriteEndElement();   // Hyperlink
        }


    }
}
