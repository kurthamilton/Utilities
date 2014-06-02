using System;
using System.Collections;
using System.Collections.Generic;

using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

using OpenXmlDrawing = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace Utilities.Office.Excel
{
    public class ShapeCollection : BaseExcel, IEnumerable<Shape>
    {
        private Drawing Drawing { get; set; }

        private SortedDictionary<int, Shape> _shapes;
        private SortedDictionary<int, Shape> Shapes { get { if (_shapes == null) { _shapes = new SortedDictionary<int, Shape>(); LoadShapes(); } return _shapes; } }

        public int Count { get { return Shapes.Count; } }
        
        /***********************************
         * CONSTRUCTORS
         ************************************/


        internal ShapeCollection(Drawing drawing)
        {
            Drawing = drawing;            
        }


        /***********************************
         * PUBLIC METHODS
         ************************************/


        public Shape this[int index]
        {
            get
            {
                if (Contains(index))
                    return Shapes[index];
                else
                    throw new ArgumentOutOfRangeException();
            }
        }

        /// <summary>
        /// Determines whether the given Shape index exists within the Drawing Shapes.
        /// </summary>
        public bool Contains(int index)
        {
            return Shapes.ContainsKey(index);
        }

        public ShapeCollection Clone(Drawing drawing)
        {
            ShapeCollection newShapes = new ShapeCollection(drawing);
            foreach (Shape shape in this)
            {
                newShapes.AddShapeToCollection(shape.Clone(drawing));
            }
            return newShapes;
        }

        // Implement IEnumerable
        public IEnumerator<Shape> GetEnumerator()
        {
            return new GenericEnumerator<Shape>(Shapes);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        /***********************************
         * PRIVATE METHODS
         ************************************/

        private void LoadShapes()
        {            
            IEnumerable<Shape> shapes = ReadShapesFromDrawing(Drawing);
            foreach (Shape shape in shapes)
            {
                AddShapeToCollection(shape);
            }
        }

        private void AddShapeToCollection(Shape shape)
        {
            if (shape != null)
                Shapes.Add(shape.Index, shape);
        }

        private void RemoveImageFromCollection(Shape shape)
        {
            if (Contains(shape.Index))
                Shapes.Remove(shape.Index);
        }


        /***********************************
         * DAL METHODS
         ************************************/

        // Read

        private static IEnumerable<Shape> ReadShapesFromDrawing(Drawing drawing)
        {
            OpenXmlPackaging.DrawingsPart drawingsPart = Drawing.GetDrawingsPartFromDrawing(drawing);
            CustomOpenXmlReader reader = CustomOpenXmlReader.Create(drawingsPart.WorksheetDrawing);
            return ReadShapesFromReader(reader, drawing);
        }

        private static IEnumerable<Shape> ReadShapesFromReader(CustomOpenXmlReader reader, Drawing drawing)
        {
            List<Shape> shapes = new List<Shape>();

            while (reader.ReadToEndElement<OpenXmlDrawing.WorksheetDrawing>())
            {
                if (reader.IsStartElementOfType<OpenXmlDrawing.AbsoluteAnchor>())
                    shapes.Add(Shape.ReadShapeFromReader<OpenXmlDrawing.AbsoluteAnchor>(reader, drawing, shapes.Count));
                else if (reader.IsStartElementOfType<OpenXmlDrawing.OneCellAnchor>())
                    shapes.Add(Shape.ReadShapeFromReader<OpenXmlDrawing.OneCellAnchor>(reader, drawing, shapes.Count));
                else if (reader.IsStartElementOfType<OpenXmlDrawing.TwoCellAnchor>())
                    shapes.Add(Shape.ReadShapeFromReader<OpenXmlDrawing.TwoCellAnchor>(reader, drawing, shapes.Count));
                    
            }

            return shapes;
        }

        // Write

    }
}
