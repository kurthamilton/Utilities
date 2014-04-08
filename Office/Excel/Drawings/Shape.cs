using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

using OpenXmlDrawing = DocumentFormat.OpenXml.Drawing;
using OpenXmlDrawingSpreadsheet = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace Utilities.Office.Excel
{

    public class Shape : BaseExcel
    {
        public Drawing Drawing { get; private set; }
        public int Index { get; private set; }        
        private ShapeAnchorType AnchorType { get; set; }
        private Anchor From { get; set; }
        private Anchor To { get; set; }
        public Picture Picture { get; private set; }
        

        /***********************************
         * CONSTRUCTORS
         ************************************/


        internal Shape(Drawing drawing, int index, ShapeAnchorType anchorType, Anchor from, Anchor to, Picture picture)
        {
            Drawing = drawing;
            Index = index;            
            AnchorType = anchorType;
            From = from;
            To = to;
            Picture = picture;
        }


        /***********************************
         * PUBLIC METHODS
         ************************************/

        public Shape Clone(Drawing drawing)
        {
            Shape newShape = new Shape(drawing, Index, AnchorType, From, To, Picture.Clone(drawing));
            return newShape;
        }

        /***********************************
         * DAL METHODS
         ************************************/

        // Read

        internal static Shape ReadShapeFromReader<T>(CustomOpenXmlReader reader, Drawing drawing, int index) where T : OpenXml.OpenXmlElement
        {
            ShapeAnchorType anchorType = GetShapeAnchorTypeFromReader(reader);
            Anchor from = new Anchor();
            Anchor to = new Anchor();
            Picture picture = null;
            
            while (reader.ReadToEndElement<T>())
            {
                if (reader.IsStartElementOfType<OpenXmlDrawingSpreadsheet.FromMarker>())
                    from = Shape.ReadAnchorFromReader<OpenXmlDrawingSpreadsheet.FromMarker>(reader);
                else if (reader.IsStartElementOfType<OpenXmlDrawingSpreadsheet.ToMarker>())
                    to = Shape.ReadAnchorFromReader<OpenXmlDrawingSpreadsheet.ToMarker>(reader);
                else if (reader.IsStartElementOfType<OpenXmlDrawingSpreadsheet.Picture>())
                    picture = Picture.ReadPictureFromReader(reader, drawing);
            }

            return new Shape(drawing, index, anchorType, from, to, picture);
        }

        private static ShapeAnchorType GetShapeAnchorTypeFromReader(CustomOpenXmlReader reader)
        {
            if (reader.ElementType == typeof(OpenXmlDrawingSpreadsheet.AbsoluteAnchor))
                return ShapeAnchorType.AbsoluteAnchor;
            else if (reader.ElementType == typeof(OpenXmlDrawingSpreadsheet.OneCellAnchor))
                return ShapeAnchorType.OneCellAnchor;
            else if (reader.ElementType == typeof(OpenXmlDrawingSpreadsheet.TwoCellAnchor))
                return ShapeAnchorType.TwoCellAnchor;
            else
                throw new Exception(string.Format("ShapeAnchorType {0} not supported in Shape.GetShapeAnchorTypeFromReader", reader.ElementType));
        }

        private static Anchor ReadAnchorFromReader<T>(CustomOpenXmlReader reader) where T : OpenXml.OpenXmlElement
        {
            Anchor anchor = new Anchor();

            while (reader.ReadToEndElement<T>())
            {
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.Column>())
                    anchor.Col = int.Parse(reader.GetText());
                else if (reader.IsStartElementOfType<OpenXmlDrawingSpreadsheet.ColumnOffset>())
                    anchor.ColOffset = int.Parse(reader.GetText());
                else if (reader.IsStartElementOfType<OpenXmlSpreadsheet.Row>())
                    anchor.Row = int.Parse(reader.GetText());
                else if (reader.IsStartElementOfType<OpenXmlDrawingSpreadsheet.RowOffset>())
                    anchor.RowOffset = int.Parse(reader.GetText());
            }

            return anchor;
        }

        // Write

        internal static void AddShapeToWorksheetDrawingElement(OpenXmlDrawingSpreadsheet.WorksheetDrawing worksheetDrawing, Shape shape)
        {
            OpenXmlDrawingSpreadsheet.AbsoluteAnchor absoluteAnchor = new OpenXmlDrawingSpreadsheet.AbsoluteAnchor();

            AddPositionToAbsoluteAnchorElement(absoluteAnchor, shape);
            AddExtentToAbsoluteAnchorElement(absoluteAnchor, shape);            
            Picture.AddPictureElementToAnchorElement(absoluteAnchor, shape.Picture);
            AddClientDataToAbsoluteAnchorElement(absoluteAnchor, shape);

            worksheetDrawing.Append(absoluteAnchor);
        }

        private static void AddPositionToAbsoluteAnchorElement(OpenXmlDrawingSpreadsheet.AbsoluteAnchor absoluteAnchor, Shape shape)
        {
            absoluteAnchor.Position = new OpenXmlDrawingSpreadsheet.Position();
            absoluteAnchor.Position.X = shape.Picture.Position.X;
            absoluteAnchor.Position.Y = shape.Picture.Position.Y;
        }

        private static void AddExtentToAbsoluteAnchorElement(OpenXmlDrawingSpreadsheet.AbsoluteAnchor absoluteAnchor, Shape shape)
        {
            absoluteAnchor.Extent = new OpenXmlDrawingSpreadsheet.Extent();
            absoluteAnchor.Extent.Cx = shape.Picture.Size.Width;
            absoluteAnchor.Extent.Cy = shape.Picture.Size.Height;
        }

        private static void AddClientDataToAbsoluteAnchorElement(OpenXmlDrawingSpreadsheet.AbsoluteAnchor absoluteAnchor, Shape shape)
        {
            absoluteAnchor.Append(new OpenXmlDrawingSpreadsheet.ClientData());
        }
    }
}
