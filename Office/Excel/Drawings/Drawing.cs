using System;
using System.Collections.Generic;
using System.Linq;

using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

using OpenXmlDrawing = DocumentFormat.OpenXml.Drawing;
using OpenXmlDrawingSpreadsheet = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace Utilities.Office.Excel
{
    public class Drawing : BaseExcel
    {        
        public Worksheet Worksheet { get; private set; }
        public string RelationshipId { get; private set; }

        private ShapeCollection _shapes;
        public ShapeCollection Shapes { get { if (_shapes == null) _shapes = new ShapeCollection(this); return _shapes; } private set { _shapes = value; } }

        /***********************************
         * CONSTRUCTORS
         ************************************/


        public Drawing(Worksheet worksheet, string relationshipId)
        {
            Worksheet = worksheet;

            if (relationshipId == "")
            {
                relationshipId = Drawing.CreateDrawingPartByDrawing(this);
            }
            RelationshipId = relationshipId;
        }


        /***********************************
         * PUBLIC METHODS
         ************************************/

        public Drawing Clone(Worksheet worksheet)
        {
            Drawing newDrawing = new Drawing(worksheet, "");           
            newDrawing.Shapes = Shapes.Clone(newDrawing);
            return newDrawing;
        }

        /// <summary>
        /// BE CAREFUL. This is only simply implemented at the moment. Images can be common to multiple worksheets, but only updated through their worksheet implementation.
        /// </summary>
        internal void Save()
        {
            Drawing.SaveDrawingData(this);
        }

        /***********************************
         * DAL METHODS
         ************************************/
        
        // Read

        internal static OpenXmlPackaging.DrawingsPart GetDrawingsPartFromDrawing(Drawing drawing)
        {
            OpenXmlPackaging.WorksheetPart worksheetPart = Worksheet.GetWorksheetPartByWorksheet(drawing.Worksheet);
            OpenXmlPackaging.DrawingsPart drawingsPart = (OpenXmlPackaging.DrawingsPart)worksheetPart.GetPartById(drawing.RelationshipId);
            return drawingsPart;
        }

        internal static Drawing ReadDrawingFromReader(CustomOpenXmlReader reader, Worksheet worksheet)
        {
            string relationshipId = "";

            foreach (CustomOpenXmlAttribute attribute in reader.Attributes)
            {
                switch (attribute.LocalName)
                {
                    case "id":
                        relationshipId = attribute.Value;
                        break;
                }
            }

            return new Drawing(worksheet, relationshipId);
        }

        // Write

        internal static void WriteDrawingToWorksheetWriter(CustomOpenXmlWriter<OpenXmlPackaging.WorksheetPart> writer, Drawing drawing)
        {
            writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.Drawing());
            writer.WriteAttribute("id", drawing.RelationshipId, "r");
            writer.WriteEndElement();   // Drawing
        }

        private static void SaveDrawingData(Drawing drawing)
        {            
            if (drawing._shapes != null)
            {
                OpenXmlPackaging.DrawingsPart drawingsPart = Drawing.GetDrawingsPartFromDrawing(drawing);

                OpenXmlDrawingSpreadsheet.WorksheetDrawing worksheetDrawing = drawingsPart.WorksheetDrawing;

                List<OpenXmlDrawingSpreadsheet.TwoCellAnchor> twoCellAnchors = new List<OpenXmlDrawingSpreadsheet.TwoCellAnchor>();
                List<OpenXmlDrawingSpreadsheet.AbsoluteAnchor> absoluteAnchors = new List<OpenXmlDrawingSpreadsheet.AbsoluteAnchor>();

                worksheetDrawing.RemoveAllChildren();

                foreach (Shape shape in drawing.Shapes)
                {
                    Shape.AddShapeToWorksheetDrawingElement(worksheetDrawing, shape);
                }

            }
        }

        private static string CreateDrawingPartByDrawing(Drawing drawing)
        {
            OpenXmlPackaging.WorkbookPart workbookPart = drawing.Worksheet.Workbook.Document.WorkbookPart;
            OpenXmlPackaging.WorksheetPart worksheetPart = Worksheet.GetWorksheetPartByWorksheet(drawing.Worksheet);
            OpenXmlPackaging.DrawingsPart drawingsPart = OpenXmlUtilities.CreatePart<OpenXmlPackaging.WorksheetPart, OpenXmlPackaging.DrawingsPart>(worksheetPart);

            using (var writer = new CustomOpenXmlWriter<OpenXmlPackaging.DrawingsPart>(drawingsPart))
            {
                writer.WriteOpenXmlElement(new OpenXmlDrawingSpreadsheet.WorksheetDrawing(), true);
            }

            string id = worksheetPart.GetIdOfPart(drawingsPart);
            worksheetPart.CreateRelationshipToPart(drawingsPart, id);
            return id;
        }
    }
}
