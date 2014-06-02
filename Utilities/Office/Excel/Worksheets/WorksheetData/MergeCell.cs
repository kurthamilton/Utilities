using System;
using System.Collections.Generic;
using System.Linq;

using System.Xml;
using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{
    internal static class MergeCell
    {

        /***********************************
         * PRIVATE METHODS
         ************************************/

        private static List<string> GetMergeCellsFromWorksheet(Worksheet worksheet)
        {
            // this method needs updating to handle MergeDown and MergeAcross together.
            List<string> mergeAreas = new List<string>();

            foreach (Row row in worksheet.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    string address = "";

                    if (cell.MergeDown > 0)
                    {
                        int currentRowIndex = row.Index;
                        // check for merge areas within the current merge area
                        for (int rowIndex = 1; rowIndex <= cell.MergeDown; rowIndex++)
                        {
                            if (worksheet.Rows.Contains(currentRowIndex + rowIndex))
                            {
                                Cell tempCell = worksheet.Rows[currentRowIndex + rowIndex].Cells[cell.Column.Index];
                                if (tempCell.MergeDown > 0)
                                {
                                    cell.MergeDown += (tempCell.MergeDown - rowIndex + 1);
                                    tempCell.MergeDown = 0;
                                }
                            }
                        }

                        address = string.Format("{0}:{1}", cell.Address, worksheet.Cells[currentRowIndex + cell.MergeDown, cell.Column.Index].Address);
                    }
                    else if (cell.MergeAcross > 0)
                    {
                        int currentColumnIndex = cell.Column.Index;
                        // check for merge areas within the current merge area
                        for (int columnIndex = 1; columnIndex < cell.MergeAcross; columnIndex++)
                        {
                            if (worksheet.Columns.Contains(currentColumnIndex + columnIndex))
                            {
                                Cell tempCell = row.Cells[currentColumnIndex + columnIndex];
                                if (tempCell.MergeAcross > 0)
                                {
                                    cell.MergeAcross += (tempCell.MergeAcross - columnIndex + 1);
                                    tempCell.MergeAcross = 0;
                                }
                            }
                        }

                        address = string.Format("{0}:{1}", cell.Address, row.Cells[currentColumnIndex + cell.MergeAcross - 1].Address);
                    }

                    if (!string.IsNullOrEmpty(address))
                        mergeAreas.Add(address);
                }
            }

            return mergeAreas;
        }

        /***********************************
         * DAL METHODS
         ************************************/

        // Read

        internal static void UpdateCellMergeCellsFromReader(CustomOpenXmlReader reader, Worksheet worksheet)
        {
            while (reader.ReadToEndElement<OpenXmlSpreadsheet.MergeCells>())
            {
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.MergeCell>())
                    UpdateCellMergeCellFromReader(reader, worksheet);
            }
        }

        private static void UpdateCellMergeCellFromReader(CustomOpenXmlReader reader, Worksheet worksheet)
        {
            // ignore reading in merged cells for the time being.

            //string range = "";

            //foreach (CustomOpenXmlAttribute attribute in reader.Attributes)
            //{
            //    switch (attribute.LocalName)
            //    {
            //        case "ref":
            //            range = attribute.Value;
            //            break;
            //    }
            //}            

            
            //string[] cellAddresses = range.Split(',');
            //Cell cellFrom = worksheet.Cells[cellAddresses[0]];
            //Cell cellTo = worksheet.Cells[cellAddresses[1]];
        }

        // Write

        internal static void WriteMergeCellsToWriter(CustomOpenXmlWriter<OpenXmlPackaging.WorksheetPart> writer, Worksheet worksheet)
        {
            List<string> mergeCells = MergeCell.GetMergeCellsFromWorksheet(worksheet);

            if (mergeCells.Count > 0)
            {
                writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.MergeCells());

                foreach (string mergeCell in mergeCells)
                {
                    MergeCell.WriteMergeCellToWriter(writer, mergeCell);
                }

                writer.WriteEndElement();   // MergeCells
            }
        }

        private static void WriteMergeCellToWriter(CustomOpenXmlWriter<OpenXmlPackaging.WorksheetPart> writer, string mergeCell)
        {
            writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.MergeCell());

            writer.WriteAttribute("ref", mergeCell);

            writer.WriteEndElement();   // MergeCell
        }


    }
}
