using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

using OpenXmlDrawing = DocumentFormat.OpenXml.Drawing;
using OpenXmlDrawingSpreadsheet = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace Utilities.Office.Excel
{

    public class Picture : BaseExcel
    {
        public Drawing Drawing { get; private set; }        

        public string FilePath { get; private set; }
        public string FileName { get { return new FileInfo(FilePath).Name; } }

        // Non-Visual Picture Properties
        public string Description { get; private set; }
        public string Name { get; private set; }
        public int Id { get; private set; }

        // Blip
        private string BlipRelationshipId { get; set; }

        public Size Size { get; internal set; }
        public Point Position { get; private set; }       

        /***********************************
         * CONSTRUCTORS
         ************************************/


        internal Picture(Drawing drawing, string filePath, string description, string name, int id, Size size, Point position, string blipRelationshipId)
        {            
            Drawing = drawing;            
            FilePath = filePath;
            Description = description;
            Name = name;
            Id = id;
            BlipRelationshipId = blipRelationshipId;
            Size = size;
            Position = position;
        }
        


        /***********************************
         * PUBLIC METHODS
         ************************************/

        public Picture Clone(Drawing drawing)
        {
            int id = Picture.GetNextPictureId(drawing.Worksheet.Workbook);
            return new Picture(drawing, FilePath, Description, Name, id, Size, Position, "");
        }

        /// <summary>
        /// A lazy method to directly replace an existing image with a new image. This will replace the instance of this image throughout the whole workbook. 
        /// This may well be deprecated at some point.
        /// </summary>
        public void ReplaceUnderlyingImageFile(string imageFilePath)
        {
            if (File.Exists(imageFilePath))
            {
                OpenXmlPackaging.OpenXmlPart mediaPart = Picture.GetMediaPartFromDrawing(Drawing, BlipRelationshipId);

                if (mediaPart != null)
                {
                    using (FileStream imageStream = File.Open(imageFilePath, FileMode.Open))
                    {
                        using (Stream mediaPartStream = mediaPart.GetStream(FileMode.Open, FileAccess.ReadWrite))
                        {
                            imageStream.CopyTo(mediaPartStream);
                            UpdatePictureSize(imageStream);
                        }
                    }
                }
            }
        }

        /***********************************
         * PRIVATE METHODS
         ************************************/

        private void UpdatePictureSize(Stream imageStream)
        {
            // Old size in EMUs
            Size oldSize = Size;

            // New size in pixels
            Size newSize = GetImageSize(imageStream);
            // Convert new size to EMUs
            newSize.Height = PixelToEmu(newSize.Height);
            newSize.Width = PixelToEmu(newSize.Width);

            if (!oldSize.Equals(newSize))
                Size = newSize;
        }

        private static Size GetImageSize(Stream imageStream)
        {
            Size size = new Size(0, 0);

            using (Image image = Image.FromStream(imageStream))
            {
                size = image.Size;
            }

            return size;
        }

        private static int PixelToEmu(int pixels)
        {
            // Open XML measurements are in EMUs (English Metric Units)
            // 12700 EMU = 1 point = 4/3px
            // 1px = 12700 * 3/4 = 9525 EMU
            return pixels * 9525;
        }

        private static int GetNextPictureId(Workbook workbook)
        {
            int maxPictureId = 0;
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                if (worksheet.Drawing != null)
                {
                    foreach (Shape shape in worksheet.Drawing.Shapes)
                    {
                        if (shape.Picture != null)
                        {
                            if (shape.Picture.Id > maxPictureId)
                                maxPictureId = shape.Picture.Id;
                        }
                    }
                }
            }
            return maxPictureId + 1;
            
        }

        /***********************************
         * DAL METHODS
         ************************************/

        // Read

        internal static Picture ReadPictureFromReader(CustomOpenXmlReader reader, Drawing drawing)
        {            
            string filePath = "";
            string description = "";
            string name = "";
            int id = 0;
            string blipRelationshipId = "";
            Size size = new Size();
            Point position = new Point();
            
            while (reader.ReadToEndElement<OpenXmlDrawingSpreadsheet.Picture>())
            {
                if (reader.IsStartElementOfType<OpenXmlDrawingSpreadsheet.NonVisualPictureProperties>())
                {
                    while (reader.ReadToEndElement<OpenXmlDrawingSpreadsheet.NonVisualPictureProperties>())
                    {
                        if (reader.IsStartElementOfType<OpenXmlDrawingSpreadsheet.NonVisualDrawingProperties>())
                        {
                            description = OpenXmlUtilities.GetAttributeValueFromReader(reader, "descr");
                            name = OpenXmlUtilities.GetAttributeValueFromReader(reader, "name");
                            id = reader.Attributes["id"].GetIntValue();
                        }                        
                    }
                }
                else if (reader.IsStartElementOfType<OpenXmlDrawingSpreadsheet.BlipFill>())
                {
                    while (reader.ReadToEndElement<OpenXmlDrawingSpreadsheet.BlipFill>())
                    {
                        if (reader.IsStartElementOfType<OpenXmlDrawing.Blip>())
                        {
                            blipRelationshipId = reader.Attributes["embed"].Value;
                            filePath = GetFilePathFromBlip(reader, drawing, blipRelationshipId);
                        }
                    }
                }
                else if (reader.IsStartElementOfType<OpenXmlDrawingSpreadsheet.ShapeProperties>())
                {
                    while (reader.ReadToEndElement<OpenXmlDrawingSpreadsheet.ShapeProperties>())
                    {
                        if (reader.IsStartElementOfType<OpenXmlDrawing.Offset>())
                            position = GetPositionFromReader(reader);
                        else if (reader.IsStartElementOfType<OpenXmlDrawing.Extents>())
                            size = GetSizeFromReader(reader);
                    }
                }
            }

            return new Picture(drawing, filePath, description, name, id, size, position, blipRelationshipId);//, sourceRectangle);
        }

        private static OpenXmlPackaging.OpenXmlPart GetMediaPartFromDrawing(Drawing drawing, string relationshipId)
        {
            OpenXmlPackaging.DrawingsPart drawingsPart = Drawing.GetDrawingsPartFromDrawing(drawing);
            OpenXmlPackaging.OpenXmlPart mediaPart = drawingsPart.GetPartById(relationshipId);
            return mediaPart;
        }

        private static string GetFilePathFromBlip(CustomOpenXmlReader reader, Drawing drawing, string relationshipId)
        {            
            OpenXmlPackaging.OpenXmlPart mediaPart = Picture.GetMediaPartFromDrawing(drawing, relationshipId);

            if (mediaPart != null)
                return mediaPart.Uri.ToString();

            return "";
        }
        
        private static Point GetPositionFromReader(CustomOpenXmlReader reader)
        {
            Point position = new Point();
            position.X = reader.Attributes["x"].GetIntValue();
            position.Y = reader.Attributes["y"].GetIntValue();
            return position;
        }

        private static Size GetSizeFromReader(CustomOpenXmlReader reader)
        {
            Size size = new Size();
            size.Width = reader.Attributes["cx"].GetIntValue();
            size.Height = reader.Attributes["cy"].GetIntValue();
            return size;
        }

        // Write

        internal static void AddPictureElementToAnchorElement(OpenXml.OpenXmlElement anchorElement, Picture picture)
        {
            OpenXmlDrawingSpreadsheet.Picture pic = new OpenXmlDrawingSpreadsheet.Picture();

            AddNonVisualPicturePropertiesToPictureElement(pic, picture);
            AddBlipFillToPictureElement(pic, picture);
            AddShapePropertiesToPictureElement(pic, picture);
            AddTransformToPictureElement(pic, picture);
            AddPresetGeometryToPictureElement(pic, picture);                  

            anchorElement.Append(pic);
        }

        private static void AddNonVisualPicturePropertiesToPictureElement(OpenXmlDrawingSpreadsheet.Picture pic, Picture picture)
        {
            pic.NonVisualPictureProperties = new OpenXmlDrawingSpreadsheet.NonVisualPictureProperties();
            var nvPicPr = pic.NonVisualPictureProperties;
            nvPicPr.NonVisualDrawingProperties = new OpenXmlDrawingSpreadsheet.NonVisualDrawingProperties();
            nvPicPr.NonVisualDrawingProperties.Description = picture.Description;
            nvPicPr.NonVisualDrawingProperties.Name = picture.Name;
            nvPicPr.NonVisualDrawingProperties.Id = (uint)picture.Id;

            nvPicPr.NonVisualPictureDrawingProperties = new OpenXmlDrawingSpreadsheet.NonVisualPictureDrawingProperties();
            var cNvPicPr = nvPicPr.NonVisualPictureDrawingProperties;
            cNvPicPr.PictureLocks = new OpenXmlDrawing.PictureLocks();
            cNvPicPr.PictureLocks.NoChangeAspect = true;
        }

        private static void AddBlipFillToPictureElement(OpenXmlDrawingSpreadsheet.Picture pic, Picture picture)
        {
            pic.BlipFill = new OpenXmlDrawingSpreadsheet.BlipFill();
            OpenXmlDrawingSpreadsheet.BlipFill blipFill = pic.BlipFill;
            blipFill.Blip = new OpenXmlDrawing.Blip();
            blipFill.Blip.CompressionState = OpenXmlDrawing.BlipCompressionValues.Print;

            OpenXmlPackaging.DrawingsPart drawingsPart = Drawing.GetDrawingsPartFromDrawing(picture.Drawing);

            if (picture.BlipRelationshipId == "")
            {
                // we know the url of the image, so create a relationship to it that way.
                OpenXmlPackaging.OpenXmlPart imagePart = OpenXmlUtilities.GetPartByUri(picture.Drawing.Worksheet.Workbook.Document.WorkbookPart, picture.FilePath);
                if (imagePart != null)
                {
                    // The workbook will probably be broken if an invalid file path is found. Not really a concern for now.                    
                    picture.BlipRelationshipId = drawingsPart.CreateRelationshipToPart(imagePart);
                }
            }

            blipFill.Blip.SetAttribute(new OpenXml.OpenXmlAttribute("r", "embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", picture.BlipRelationshipId));

            var stretch = new OpenXmlDrawing.Stretch();
            stretch.FillRectangle = new OpenXmlDrawing.FillRectangle();
            blipFill.Append(stretch);
        }

        private static void AddShapePropertiesToPictureElement(OpenXmlDrawingSpreadsheet.Picture pic, Picture picture)
        {
            pic.ShapeProperties = new OpenXmlDrawingSpreadsheet.ShapeProperties();
            pic.ShapeProperties.BlackWhiteMode = OpenXmlDrawing.BlackWhiteModeValues.Auto;
        }

        private static void AddTransformToPictureElement(OpenXmlDrawingSpreadsheet.Picture pic, Picture picture)
        {
            var xfrm = new OpenXmlDrawing.Transform2D();

            xfrm.Offset = new OpenXmlDrawing.Offset();
            xfrm.Offset.X = picture.Position.X;
            xfrm.Offset.Y = picture.Position.Y;

            xfrm.Extents = new OpenXmlDrawing.Extents();
            xfrm.Extents.Cx = picture.Size.Width;
            xfrm.Extents.Cy = picture.Size.Height;

            pic.ShapeProperties.Transform2D = xfrm;            
        }

        private static void AddPresetGeometryToPictureElement(OpenXmlDrawingSpreadsheet.Picture pic, Picture picture)
        {
            OpenXmlDrawing.PresetGeometry prstGeom = new OpenXmlDrawing.PresetGeometry();
            prstGeom.Preset = OpenXmlDrawing.ShapeTypeValues.Rectangle;            
            pic.ShapeProperties.Append(prstGeom);
        }
    }
}
