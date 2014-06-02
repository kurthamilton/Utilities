using System;
using System.Collections.Generic;
using SystemDrawing = System.Drawing;

using System.Xml;
using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{    

    public class Color : BaseExcel, IEquatable<Color>
    {
        public event EventHandler Updated;

        /// <summary>
        /// Indexed is deprecated in Open XML, but carried over for legacy support.
        /// </summary>
        public int Indexed 
        { 
            get { return _indexed; }
            set { if (value != _indexed) { if (value >= 0) ResetProperties(); _indexed = value; UpdateProperty(); } } 
        }
        private int _indexed;

        public int Theme 
        { 
            get { return _theme; }
            set { if (value != _theme) { if (value >= 0) ResetProperties(); _theme = value; UpdateProperty(); } } 
        }
        private int _theme;

        /// <summary>
        /// Html RGB value. Set with Color.GetRgb(r,g,b), or a #RRGGBB hex code.
        /// </summary>
        public string Rgb 
        { 
            get { return _rgb; }
            set { if (value != _rgb) { if (!string.IsNullOrEmpty(value)) ResetProperties(); _rgb = value; UpdateProperty(); } } 
        }
        private string _rgb;

        // not looked into how tint is used in conjuction with other properties, so leave internal for now.
        internal double Tint
        {
            get { return _tint; }
            set { _tint = value; UpdateProperty(); }
        }
        private double _tint;

        public bool Auto
        {
            get { return _auto; }
            set { if (value) ResetProperties(); _auto = value; UpdateProperty(); }
        }
        private bool _auto;

        //private BaseRange BaseRange { get; set; }

        /***********************************
         * CONSTRUCTORS
         ************************************/

        internal Color(BaseRange range)
        {
            //BaseRange = range;
        }
        internal Color(BaseRange range, Color color)
            : this(color)
        {
            //BaseRange = range;
        }

        public Color()
        {
            ResetProperties();
        }

        internal Color(Color color)
            : this(color.Indexed, color.Theme, color.Rgb, color.Tint, color.Auto)
        {            
        }        

        internal Color(int indexed, int theme, string rgb, double tint, bool auto)
        {
            Indexed = indexed;
            Theme = theme;
            Rgb = rgb;
            Tint = tint;
            Auto = auto;
        }

        /***********************************
         * PUBLIC METHODS
         ************************************/

        public static string GetRgb(int red, int green, int blue)
        {
            return SystemDrawing.ColorTranslator.ToHtml(SystemDrawing.Color.FromArgb(red, green, blue));
        }

        public Color Clone()
        {
            Color clone = new Color();
            clone.Auto = Auto;
            clone.Indexed = Indexed;
            clone.Rgb = Rgb;
            clone.Theme = Theme;
            clone.Tint = Tint;
            return clone;
        }

        // implement IEquatable
        public bool Equals(Color other)
        {
            bool isUsed = IsUsed();
            bool otherIsUsed = other.IsUsed();

            if (isUsed && otherIsUsed)
            {
                return (other.Indexed == Indexed &&
                    other.Theme == Theme &&
                    other.Rgb == Rgb && 
                    other.Tint == Tint && 
                    other.Auto == Auto);
            }
            else
                return (isUsed == otherIsUsed);
        }

        /***********************************
         * INTERNAL METHODS
         ************************************/

        internal bool IsUsed()
        {
            return (_indexed > -1 || _theme > -1 || !string.IsNullOrEmpty(_rgb) || _tint != 0 || _auto == true);
        }

        /***********************************
         * PRIVATE METHODS
         ************************************/
        private void ResetProperties()
        {
            // Prevent any ambiguity between which property is being used to source the colour, so reset them with each positive property update.
            _indexed = -1;
            _theme = -1;
            _rgb = null;
            _tint = 0;
            _auto = false;
        }        

        private void UpdateProperty()
        {
            if (Updated != null)
                Updated.Invoke(this, new EventArgs());
        }

        private static bool ValidateRgb(string rgb)
        {
            if (!string.IsNullOrEmpty(rgb))
            {
                if (rgb.StartsWith("#"))
                {
                    if (rgb.Length == 7)
                    {
                        // also need to validate hex values themselves
                        return true;
                    }
                }
            }
            return false;
        }

        private static string ConvertRgbToHtml(string rgb)
        {
            // xml rgb values seem to be stored as FF + 6-digit html. Use the #RRGGBB code here for familiarity.

            if (!string.IsNullOrEmpty(rgb))
                return "#" + rgb.Substring(2);
            else
                return rgb;
        }
        private static string ConvertHtmlToRgb(string html)
        {
            if (!string.IsNullOrEmpty(html))
                return "FF" + html.Replace("#", "");
            else 
                return html;
        }


        /***********************************
         * DAL METHODS
         ************************************/

        // Read

        internal static Color ReadColorFromReader(CustomOpenXmlReader reader)
        {
            Color color = new Color();

            foreach (CustomOpenXmlAttribute attribute in reader.Attributes)
            {
                switch (attribute.LocalName)
                {
                    case "indexed":
                        color.Indexed = attribute.GetIntValue();
                        break;
                    case "theme":
                        color.Theme = attribute.GetIntValue();
                        break;
                    case "rgb":
                        color.Rgb = Color.ConvertRgbToHtml(attribute.Value);
                        break;
                    case "tint":
                        color.Tint = attribute.GetDoubleValue();
                        break;
                    case "auto":
                        color.Auto = attribute.GetBoolValue();
                        break;
                }
            }

            return color;
        }

        internal static List<string> ReadIndexedColorsFromReader(CustomOpenXmlReader reader)
        {
            List<string> indexedColors = new List<string>();

            while (reader.ReadToEndElement<OpenXmlSpreadsheet.IndexedColors>())
            {
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.RgbColor>())
                {
                    indexedColors.Add(reader.Attributes["rgb"].Value);
                }
            }

            return indexedColors;
        }

        // Write

        internal static void WriteColorToWriter<T>(CustomOpenXmlWriter<T> writer, Color color, OpenXml.OpenXmlElement colorElement = null) where T : OpenXmlPackaging.OpenXmlPart, OpenXmlPackaging.IFixedContentTypePart
        {
            if (colorElement == null)
                colorElement = new OpenXmlSpreadsheet.Color();
            writer.WriteOpenXmlElement(colorElement);

            if (color.Indexed >= 0) writer.WriteAttribute("indexed", color.Indexed);
            if (color.Theme >= 0) writer.WriteAttribute("theme", color.Theme);
            if (!string.IsNullOrEmpty(color.Rgb)) writer.WriteAttribute("rgb", Color.ConvertHtmlToRgb(color.Rgb));
            if (color.Tint != 0) writer.WriteAttribute("tint", color.Tint);
            if (color.Auto) writer.WriteAttribute("auto", color.Auto);
            
            writer.WriteEndElement();   // colorElement
        }

        internal static void WriteIndexedColorsToWriter(CustomOpenXmlWriter<OpenXmlPackaging.WorkbookStylesPart> writer, List<string> indexedColors)
        {
            writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.Colors());
            writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.IndexedColors());

            foreach (string indexedColor in indexedColors)
            {                
                writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.RgbColor());
                writer.WriteAttribute("rgb", indexedColor);
                writer.WriteEndElement();   // RgbColor
            }

            writer.WriteEndElement();   // IndexedColors
            writer.WriteEndElement();   // Colors
        }
    }
}
